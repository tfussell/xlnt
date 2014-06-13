#include <algorithm>
#include <pugixml.hpp>

#include "reader/reader.hpp"
#include "cell/cell.hpp"
#include "common/datetime.hpp"
#include "worksheet/range_reference.hpp"
#include "workbook/workbook.hpp"
#include "worksheet/worksheet.hpp"
#include "workbook/document_properties.hpp"
#include "common/zip_file.hpp"

namespace xlnt {

std::vector<std::pair<std::string, std::string>> reader::read_sheets(const zip_file &archive)
{
    auto xml_source = archive.get_file_contents("xl/workbook.xml");
    pugi::xml_document doc;
    doc.load(xml_source.c_str());
    std::vector<std::pair<std::string, std::string>> sheets;
    for(auto sheet_node : doc.child("workbook").child("sheets").children("sheet"))
    {
        std::string id = sheet_node.attribute("r:id").as_string();
        std::string name = sheet_node.attribute("name").as_string();
        sheets.push_back(std::make_pair(id, name));
    }
    return sheets;
}

document_properties reader::read_properties_core(const std::string &xml_string)
{
    document_properties props;
    pugi::xml_document doc;
    doc.load(xml_string.c_str());
    auto root_node = doc.child("cp:coreProperties");
    props.creator = root_node.child("dc:creator").text().as_string();
    return props;
}

std::string reader::read_dimension(const std::string &xml_string)
{
    pugi::xml_document doc;
    doc.load(xml_string.c_str());
    auto root_node = doc.child("worksheet");
    auto dimension_node = root_node.child("dimension");
    std::string dimension = dimension_node.attribute("ref").as_string();
    return dimension;
}
    
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
        relationships[id] = std::make_pair(target, type);
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
                                           { return v.second.second == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"; });
    std::string type;
    
    if(relationship_match != root_relationships.end())
    {
        std::string office_document_relationship = relationship_match->second.first;
        
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
		    date date(1900, 1, 1);
                    int days = cell_node.child("v").text().as_int();
                    while(days > 365)
                    {
                        date.year += 1;
                        days -= 365;
                    }
                    while(days > 30)
                    {
                        date.month += 1;
                        days -= 30;
                    }
                    date.day = days;
                    ws.get_cell(address) = date;
                }
                else if(cell_node.attribute("s") != nullptr && std::string(cell_node.attribute("s").as_string()) == "3") // time
                {
		    time time;
                    double remaining = cell_node.child("v").text().as_double() * 24;
                    time.hour = (int)(remaining);
                    remaining -= time.hour;
                    remaining *= 60;
                    time.minute = (int)(remaining);
                    remaining -= time.minute;
                    remaining *= 60;
                    time.second = (int)(remaining);
                    remaining -= time.second;
                    if(remaining > 0.5)
                    {
                        time.second++;
                        if(time.second > 59)
                        {
                            time.second = 0;
                            time.minute++;
                            if(time.minute > 59)
                            {
                                time.minute = 0;
                                time.hour++;
                            }
                        }
                    }
                    ws.get_cell(address) = time;
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
