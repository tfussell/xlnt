#include <algorithm>
#include <pugixml.hpp>

#include "reader.h"
#include "cell.h"
#include "range_reference.h"
#include "workbook.h"
#include "worksheet.h"

namespace xlnt {

std::unordered_map<std::string, std::pair<std::string, std::string>> reader::read_relationships(const std::string &content)
{
    pugi::xml_document doc;
    doc.load(content.c_str());
    
    auto root_node = doc.child("Relationships");
    
    std::unordered_map<std::string, std::pair<std::string, std::string>> relationships;
    
    for(auto relationship : root_node.children("Relationship"))
    {
        std::string id = relationship.attribute("Id").as_string();
        std::string type = relationship.attribute("Type").as_string();
        std::string target = relationship.attribute("Target").as_string();
        relationships[id] = std::make_pair(type, target);
    }
    
    return relationships;
}

std::pair<std::unordered_map<std::string, std::string>, std::unordered_map<std::string, std::string>> reader::read_content_types(const std::string &content)
{
    pugi::xml_document doc;
    doc.load(content.c_str());
    
    auto root_node = doc.child("Types");
    
    std::unordered_map<std::string, std::string> default_types;
    
    for(auto child : root_node.children("Default"))
    {
        default_types[child.attribute("Extension").as_string()] = child.attribute("ContentType").as_string();
    }
    
    std::unordered_map<std::string, std::string> override_types;
    
    for(auto child : root_node.children("Override"))
    {
        override_types[child.attribute("PartName").as_string()] = child.attribute("ContentType").as_string();
    }
    
    return std::make_pair(default_types, override_types);
}
    
std::string reader::determine_document_type(const std::unordered_map<std::string, std::pair<std::string, std::string>> &root_relationships,
                                    const std::unordered_map<std::string, std::string> &override_types)
{
    auto relationship_match = std::find_if(root_relationships.begin(), root_relationships.end(),
                                           [](const std::pair<std::string, std::pair<std::string, std::string>> &v)
                                           { return v.second.first == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"; });
    std::string type;
    
    if(relationship_match != root_relationships.end())
    {
        std::string office_document_relationship = relationship_match->second.second;
        
        if(office_document_relationship[0] != '/')
        {
            office_document_relationship = std::string("/") + office_document_relationship;
        }
        
        type = override_types.at(office_document_relationship);
    }
    
    if(type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
    {
        return "excel";
    }
    else if(type == "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml")
    {
        return "powerpoint";
    }
    else if(type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")
    {
        return "word";
    }
    
    return "unsupported";
}
    
void read_worksheet_common(worksheet ws, const pugi::xml_node &root_node, const std::vector<std::string> &string_table)
{
    auto dimension_node = root_node.child("dimension");
    std::string dimension = dimension_node.attribute("ref").as_string();
    auto sheet_data_node = root_node.child("sheetData");
    auto merge_cells_node = root_node.child("mergeCells");
    
    if(merge_cells_node != nullptr)
    {
        int count = merge_cells_node.attribute("count").as_int();
        
        for(auto merge_cell_node : merge_cells_node.children("mergeCell"))
        {
            ws.merge_cells(merge_cell_node.attribute("ref").as_string());
            count--;
        }
        
        if(count != 0)
        {
            throw std::runtime_error("mismatch between count and actual number of merged cells");
        }
    }
    
    for(auto row_node : sheet_data_node.children("row"))
    {
        int row_index = row_node.attribute("r").as_int();
        std::string span_string = row_node.attribute("spans").as_string();
        auto colon_index = span_string.find(':');
        int min_column = std::stoi(span_string.substr(0, colon_index));
        int max_column = std::stoi(span_string.substr(colon_index + 1));
        
        for(int i = min_column; i < max_column + 1; i++)
        {
            std::string address = xlnt::cell_reference::column_string_from_index(i) + std::to_string(row_index);
            auto cell_node = row_node.find_child_by_attribute("c", "r", address.c_str());
            
            if(cell_node != nullptr)
            {
                if(cell_node.attribute("t") != nullptr && std::string(cell_node.attribute("t").as_string()) == "inlineStr") // inline string
                {
                    ws.get_cell(address) = cell_node.child("is").child("t").text().as_string();
                }
                else if(cell_node.attribute("t") != nullptr && std::string(cell_node.attribute("t").as_string()) == "s") // shared string
                {
                    ws.get_cell(address) = string_table.at(cell_node.child("v").text().as_int());
                }
                else if(cell_node.attribute("s") != nullptr && std::string(cell_node.attribute("s").as_string()) == "2") // date
                {
                    tm date = tm();
                    date.tm_year = 1900;
                    int days = cell_node.child("v").text().as_int();
                    while(days > 365)
                    {
                        date.tm_year += 1;
                        days -= 365;
                    }
                    date.tm_yday = days;
                    while(days > 30)
                    {
                        date.tm_mon += 1;
                        days -= 30;
                    }
                    date.tm_mday = days;
                    ws.get_cell(address) = date;
                }
                else if(cell_node.attribute("s") != nullptr && std::string(cell_node.attribute("s").as_string()) == "3") // time
                {
                    tm date;
                    double remaining = cell_node.child("v").text().as_double() * 24;
                    date.tm_hour = (int)(remaining);
                    remaining -= date.tm_hour;
                    remaining *= 60;
                    date.tm_min = (int)(remaining);
                    remaining -= date.tm_min;
                    remaining *= 60;
                    date.tm_sec = (int)(remaining);
                    remaining -= date.tm_sec;
                    if(remaining > 0.5)
                    {
                        date.tm_sec++;
                        if(date.tm_sec > 59)
                        {
                            date.tm_sec = 0;
                            date.tm_min++;
                            if(date.tm_min > 59)
                            {
                                date.tm_min = 0;
                                date.tm_hour++;
                            }
                        }
                    }
                    ws.get_cell(address) = date;
                }
                else if(cell_node.attribute("s") != nullptr && std::string(cell_node.attribute("s").as_string()) == "4") // decimal
                {
                    ws.get_cell(address) = cell_node.child("v").text().as_double();
                }
                else if(cell_node.attribute("s") != nullptr && std::string(cell_node.attribute("s").as_string()) == "1") // percent
                {
                    ws.get_cell(address) = cell_node.child("v").text().as_double();
                }
                else if(cell_node.child("v") != nullptr)
                {
                    ws.get_cell(address) = cell_node.child("v").text().as_string();
                }
            }
        }
    }
}

void reader::read_worksheet(worksheet ws, const std::string &xml_string, const std::vector<std::string> &string_table)
{
    pugi::xml_document doc;
    doc.load(xml_string.c_str());
    read_worksheet_common(ws, doc.child("worksheet"), string_table);
}

worksheet xlnt::reader::read_worksheet(std::istream &handle, xlnt::workbook &wb, const std::string &title, const std::vector<std::string> &string_table)
{
    auto ws = wb.create_sheet();
    ws.set_title(title);
    pugi::xml_document doc;
    doc.load(handle);
    read_worksheet_common(ws, doc.child("worksheet"), string_table);
    return ws;
}

std::vector<std::string> reader::read_shared_string(const std::string &xml_string)
{
    std::vector<std::string> shared_strings;
    pugi::xml_document doc;
    doc.load(xml_string.c_str());
    auto root_node = doc.child("sst");
    //int count = root_node.attribute("count").as_int();
    int unique_count = root_node.attribute("uniqueCount").as_int();
    
    for(auto si_node : root_node)
    {
        shared_strings.push_back(si_node.child("t").text().as_string());
    }
    
    if(unique_count != (int)shared_strings.size())
    {
        throw std::runtime_error("counts don't match");
    }
    
    return shared_strings;
}
    
} // namespace xlnt
