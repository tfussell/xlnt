#include <xlnt/s11n/worksheet_serializer.hpp>
#include <xlnt/cell/cell.hpp>
#include <xlnt/cell/cell_reference.hpp>
#include <xlnt/s11n/xml_document.hpp>
#include <xlnt/s11n/xml_node.hpp>
#include <xlnt/worksheet/range.hpp>
#include <xlnt/worksheet/range_reference.hpp>

#include "detail/constants.hpp"

namespace {

bool is_integral(long double d)
{
    return d == static_cast<long long int>(d);
}
    
} // namepsace

namespace xlnt {

bool worksheet_serializer::write_worksheet(const worksheet ws, const std::vector<std::string> &string_table, xml_document &xml)
{
    ws.get_cell("A1");
    
    xml.add_namespace("", constants::Namespaces.at("spreadsheetml"));
    xml.add_namespace("r", constants::Namespaces.at("r"));

    auto &root_node = xml.root();
    root_node.set_name("worksheet");
    
    auto &sheet_pr_node = root_node.add_child("sheetPr");
    
    if(!ws.get_page_setup().is_default())
    {
        auto &page_set_up_pr_node = sheet_pr_node.add_child("pageSetUpPr");
        page_set_up_pr_node.add_attribute("fitToPage", ws.get_page_setup().fit_to_page() ? "1" : "0");
    }
    
    auto &outline_pr_node = sheet_pr_node.add_child("outlinePr");
    
    outline_pr_node.add_attribute("summaryBelow", "1");
    outline_pr_node.add_attribute("summaryRight", "1");
    
    auto &dimension_node = root_node.add_child("dimension");
    dimension_node.add_attribute("ref", ws.calculate_dimension().to_string());
    
    auto &sheet_views_node = root_node.add_child("sheetViews");
    auto &sheet_view_node = sheet_views_node.add_child("sheetView");
    sheet_view_node.add_attribute("workbookViewId", "0");
    
    std::string active_pane = "bottomRight";
    
    if(ws.has_frozen_panes())
    {
        auto pane_node = sheet_view_node.add_child("pane");
        
        if(ws.get_frozen_panes().get_column_index() > 1)
        {
            pane_node.add_attribute("xSplit", std::to_string(ws.get_frozen_panes().get_column_index() - 1));
            active_pane = "topRight";
        }
        
        if(ws.get_frozen_panes().get_row() > 1)
        {
            pane_node.add_attribute("ySplit", std::to_string(ws.get_frozen_panes().get_row() - 1));
            active_pane = "bottomLeft";
        }
        
        if(ws.get_frozen_panes().get_row() > 1 && ws.get_frozen_panes().get_column_index() > 1)
        {
            auto top_right_node = sheet_view_node.add_child("selection");
            top_right_node.add_attribute("pane", "topRight");
            auto bottom_left_node = sheet_view_node.add_child("selection");
            bottom_left_node.add_attribute("pane", "bottomLeft");
            active_pane = "bottomRight";
        }
        
        pane_node.add_attribute("topLeftCell", ws.get_frozen_panes().to_string());
        pane_node.add_attribute("activePane", active_pane);
        pane_node.add_attribute("state", "frozen");
    }
    
    auto selection_node = sheet_view_node.add_child("selection");
    
    if(ws.has_frozen_panes())
    {
        if(ws.get_frozen_panes().get_row() > 1 && ws.get_frozen_panes().get_column_index() > 1)
        {
            selection_node.add_attribute("pane", "bottomRight");
        }
        else if(ws.get_frozen_panes().get_row() > 1)
        {
            selection_node.add_attribute("pane", "bottomLeft");
        }
        else if(ws.get_frozen_panes().get_column_index() > 1)
        {
            selection_node.add_attribute("pane", "topRight");
        }
    }
    
    std::string active_cell = "A1";
    selection_node.add_attribute("activeCell", active_cell);
    selection_node.add_attribute("sqref", active_cell);
    
    auto sheet_format_pr_node = root_node.add_child("sheetFormatPr");
    sheet_format_pr_node.add_attribute("baseColWidth", "10");
    sheet_format_pr_node.add_attribute("defaultRowHeight", "15");
    
    bool has_column_properties = false;
    
    for(auto column = ws.get_lowest_column(); column <= ws.get_highest_column(); column++)
    {
        if(ws.has_column_properties(column))
        {
            has_column_properties = true;
            break;
        }
    }
    
    if(has_column_properties)
    {
        auto cols_node = root_node.add_child("cols");
        
        for(auto column = ws.get_lowest_column(); column <= ws.get_highest_column(); column++)
        {
            const auto &props = ws.get_column_properties(column);
            
            auto col_node = cols_node.add_child("col");
            
            col_node.add_attribute("min", std::to_string(column));
            col_node.add_attribute("max", std::to_string(column));
            col_node.add_attribute("width", std::to_string(props.width));
            col_node.add_attribute("style", std::to_string(props.style));
            col_node.add_attribute("customWidth", props.custom ? "1" : "0");
        }
    }
    
    std::unordered_map<std::string, std::string> hyperlink_references;
    
    auto sheet_data_node = root_node.add_child("sheetData");
    
    for(auto row : ws.rows())
    {
        row_t min = static_cast<row_t>(row.num_cells());
        row_t max = 0;
        bool any_non_null = false;
        
        for(auto cell : row)
        {
            min = std::min(min, cell_reference::column_index_from_string(cell.get_column()));
            max = std::max(max, cell_reference::column_index_from_string(cell.get_column()));
            
            if(!cell.garbage_collectible())
            {
                any_non_null = true;
            }
        }
        
        if(!any_non_null)
        {
            continue;
        }
        
        auto row_node = sheet_data_node.add_child("row");
        
        row_node.add_attribute("r", std::to_string(row.front().get_row()));
        row_node.add_attribute("spans", (std::to_string(min) + ":" + std::to_string(max)));
        
        if(ws.has_row_properties(row.front().get_row()))
        {
            row_node.add_attribute("customHeight", "1");
            auto height = ws.get_row_properties(row.front().get_row()).height;
            
            if(height == std::floor(height))
            {
                row_node.add_attribute("ht", std::to_string(static_cast<long long int>(height)) + ".0");
            }
            else
            {
                row_node.add_attribute("ht", std::to_string(height));
            }
        }
        
        //row_node.add_attribute("x14ac:dyDescent", 0.25);
        
        for(auto cell : row)
        {
            if(!cell.garbage_collectible())
            {
                if(cell.has_hyperlink())
                {
                    hyperlink_references[cell.get_hyperlink().get_id()] = cell.get_reference().to_string();
                }
                
                auto cell_node = row_node.add_child("c");
                cell_node.add_attribute("r", cell.get_reference().to_string());
                
                if(cell.get_data_type() == cell::type::string)
                {
                    if(cell.has_formula())
                    {
                        cell_node.add_attribute("t", "str");
                        cell_node.add_child("f").set_text(cell.get_formula());
                        cell_node.add_child("v").set_text(cell.to_string());
                        
                        continue;
                    }
                    
                    int match_index = -1;
                    
                    for(std::size_t i = 0; i < string_table.size(); i++)
                    {
                        if(string_table[i] == cell.get_value<std::string>())
                        {
                            match_index = static_cast<int>(i);
                            break;
                        }
                    }
                    
                    if(match_index == -1)
                    {
                        if(cell.get_value<std::string>().empty())
                        {
                            cell_node.add_attribute("t", "s");
                        }
                        else
                        {
                            cell_node.add_attribute("t", "inlineStr");
                            auto inline_string_node = cell_node.add_child("is");
                            inline_string_node.add_child("t").text().set(cell.get_value<std::string>());
                        }
                    }
                    else
                    {
                        cell_node.add_attribute("t", "s");
                        auto value_node = cell_node.add_child("v");
                        value_node.text().set(match_index);
                    }
                }
                else
                {
                    if(cell.get_data_type() != cell::type::null)
                    {
                        if(cell.get_data_type() == cell::type::boolean)
                        {
                            cell_node.add_attribute("t", "b");
                            auto value_node = cell_node.add_child("v");
                            value_node.text().set(cell.get_value<bool>() ? 1 : 0);
                        }
                        else if(cell.get_data_type() == cell::type::numeric)
                        {
                            if(cell.has_formula())
                            {
                                cell_node.add_child("f").text().set(cell.get_formula());
                                cell_node.add_child("v").text().set(cell.to_string());
                                continue;
                            }
                            
                            cell_node.add_attribute("t", "n");
                            auto value_node = cell_node.add_child("v");
                            if(is_integral(cell.get_value<long double>()))
                            {
                                value_node.text().set(cell.get_value<long long>());
                            }
                            else
                            {
                                std::stringstream ss;
                                ss.precision(20);
                                ss << cell.get_value<long double>();
                                ss.str();
                                value_node.text().set(ss.str());
                            }
                        }
                    }
                    else if(cell.has_formula())
                    {
                        cell_node.add_child("f").text().set(cell.get_formula());
                        cell_node.add_child("v");
                        continue;
                    }
                }
                
                //if(cell.has_style())
                {
                    auto style_id = cell.get_style_id();
                    cell_node.add_attribute("s", (int)style_id);
                }
            }
        }
    }
    
    if(ws.has_auto_filter())
    {
        auto auto_filter_node = root_node.add_child("autoFilter");
        auto_filter_node.add_attribute("ref", ws.get_auto_filter().to_string());
    }
    
    if(!ws.get_merged_ranges().empty())
    {
        auto merge_cells_node = root_node.add_child("mergeCells");
        merge_cells_node.add_attribute("count", (unsigned int)ws.get_merged_ranges().size());
        
        for(auto merged_range : ws.get_merged_ranges())
        {
            auto merge_cell_node = merge_cells_node.add_child("mergeCell");
            merge_cell_node.add_attribute("ref", merged_range.to_string());
        }
    }
    
    if(!ws.get_relationships().empty())
    {
        auto hyperlinks_node = root_node.add_child("hyperlinks");
        
        for(auto relationship : ws.get_relationships())
        {
            auto hyperlink_node = hyperlinks_node.add_child("hyperlink");
            hyperlink_node.add_attribute("display", relationship.get_target_uri());
            hyperlink_node.add_attribute("ref", hyperlink_references.at(relationship.get_id()));
            hyperlink_node.add_attribute("r:id", relationship.get_id());
        }
    }
    
    if(!ws.get_page_setup().is_default())
    {
        auto print_options_node = root_node.add_child("printOptions");
        print_options_node.add_attribute("horizontalCentered", ws.get_page_setup().get_horizontal_centered() ? 1 : 0);
        print_options_node.add_attribute("verticalCentered", ws.get_page_setup().get_vertical_centered() ? 1 : 0);
    }
    
    auto page_margins_node = root_node.add_child("pageMargins");
    
    page_margins_node.add_attribute("left", ws.get_page_margins().get_left());
    page_margins_node.add_attribute("right", ws.get_page_margins().get_right());
    page_margins_node.add_attribute("top", ws.get_page_margins().get_top());
    page_margins_node.add_attribute("bottom", ws.get_page_margins().get_bottom());
    page_margins_node.add_attribute("header", ws.get_page_margins().get_header());
    page_margins_node.add_attribute("footer", ws.get_page_margins().get_footer());
    
    if(!ws.get_page_setup().is_default())
    {
        auto page_setup_node = root_node.add_child("pageSetup");
        
        std::string orientation_string = ws.get_page_setup().get_orientation() == page_setup::orientation::landscape ? "landscape" : "portrait";
        page_setup_node.add_attribute("orientation", orientation_string);
        page_setup_node.add_attribute("paperSize", (int)ws.get_page_setup().get_paper_size());
        page_setup_node.add_attribute("fitToHeight", ws.get_page_setup().fit_to_height() ? 1 : 0);
        page_setup_node.add_attribute("fitToWidth", ws.get_page_setup().fit_to_width() ? 1 : 0);
    }
    
    if(!ws.get_header_footer().is_default())
    {
        auto header_footer_node = root_node.add_child("headerFooter");
        auto odd_header_node = header_footer_node.add_child("oddHeader");
        std::string header_text = "&L&\"Calibri,Regular\"&K000000Left Header Text&C&\"Arial,Regular\"&6&K445566Center Header Text&R&\"Arial,Bold\"&8&K112233Right Header Text";
        odd_header_node.text().set(header_text);
        auto odd_footer_node = header_footer_node.add_child("oddFooter");
        std::string footer_text = "&L&\"Times New Roman,Regular\"&10&K445566Left Footer Text_x000D_And &D and &T&C&\"Times New Roman,Bold\"&12&K778899Center Footer Text &Z&F on &A&R&\"Times New Roman,Italic\"&14&KAABBCCRight Footer Text &P of &N";
        odd_footer_node.text().set(footer_text);
    }
    
    std::stringstream ss;
    doc.save(ss);
    
    return ss.str();
}

} // namespace xlnt
