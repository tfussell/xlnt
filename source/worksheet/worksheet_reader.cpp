#include <xlnt/cell/cell.hpp>
#include <xlnt/cell/cell_reference.hpp>
#include <xlnt/common/zip_file.hpp>
#include <xlnt/reader/worksheet_reader.hpp>
#include <xlnt/styles/number_format.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/worksheet/range_reference.hpp>
#include <xlnt/worksheet/worksheet.hpp>

#include "detail/include_pugixml.hpp"

namespace {
    
} // namespace

namespace xlnt {

void read_worksheet(worksheet ws, zip_file &archive, const relationship &rel, const std::vector<std::string> &string_table)
{    
    pugi::xml_document doc;
    doc.load(archive.read(rel.get_target_uri()).c_str());
    
    auto root_node = doc.child("worksheet");
    
    auto dimension_node = root_node.child("dimension");
    std::string dimension = dimension_node.attribute("ref").as_string();
    auto full_range = xlnt::range_reference(dimension);
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
        
        column_t min_column = 0;
        column_t max_column = 0;
        
        if(colon_index != std::string::npos)
        {
            min_column = static_cast<column_t>(std::stoll(span_string.substr(0, colon_index)));
            max_column = static_cast<column_t>(std::stoll(span_string.substr(colon_index + 1)));
        }
        else
        {
            min_column = static_cast<column_t>(full_range.get_top_left().get_column_index());
            max_column = static_cast<column_t>(full_range.get_bottom_right().get_column_index());
        }
        
        for(column_t i = min_column; i <= max_column; i++)
        {
            std::string address = xlnt::cell_reference::column_string_from_index(i) + std::to_string(row_index);
            auto cell_node = row_node.find_child_by_attribute("c", "r", address.c_str());
            
            if(cell_node != nullptr)
            {
                bool has_value = cell_node.child("v") != nullptr;
                std::string value_string = has_value ? cell_node.child("v").text().as_string() : "";
                
                bool has_type = cell_node.attribute("t") != nullptr;
                std::string type = has_type ? cell_node.attribute("t").as_string() : "";
                
                bool has_style = cell_node.attribute("s") != nullptr;
                int style_id = has_style ? cell_node.attribute("s").as_int() : 0;
                
                bool has_formula = cell_node.child("f") != nullptr;
                bool has_shared_formula = has_formula && cell_node.child("f").attribute("t") != nullptr && std::string(cell_node.child("f").attribute("t").as_string()) == "shared";
                
                auto cell = ws.get_cell(address);
                
                if(has_formula && !has_shared_formula && !ws.get_parent().get_data_only())
                {
                    std::string formula = cell_node.child("f").text().as_string();
                    cell.set_formula(formula);
                }
                
                if(has_type && type == "inlineStr") // inline string
                {
                    std::string inline_string = cell_node.child("is").child("t").text().as_string();
                    cell.set_value(inline_string);
                }
                else if(has_type && type == "s" && !has_formula) // shared string
                {
                    auto shared_string_index = std::stoull(value_string);
                    auto shared_string = string_table.at(shared_string_index);
                    cell.set_value(shared_string);
                }
                else if(has_type && type == "b") // boolean
                {
                    cell.set_value(value_string != "0");
                }
                else if(has_type && type == "str")
                {
                    cell.set_value(value_string);
                }
                else if(has_value && !value_string.empty())
                {
                    if(!value_string.empty() && value_string[0] == '#')
                    {
                        cell.set_error(value_string);
                    }
                    else
                    {
                        cell.set_value(std::stold(value_string));
                    }
                }
                
                if(has_style)
                {
                    cell.set_style_id(style_id);
                }
            }
        }
    }
    
    auto auto_filter_node = root_node.child("autoFilter");
    
    if(auto_filter_node != nullptr)
    {
        xlnt::range_reference ref(auto_filter_node.attribute("ref").as_string());
        ws.auto_filter(ref);
    }
}
    
} // namespace xlnt
