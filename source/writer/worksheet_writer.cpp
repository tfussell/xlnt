#include <algorithm>
#include <array>
#include <sstream>

#include <xlnt/cell/cell.hpp>
#include <xlnt/cell/cell_reference.hpp>
#include <xlnt/styles/border.hpp>
#include <xlnt/styles/fill.hpp>
#include <xlnt/styles/font.hpp>
#include <xlnt/styles/number_format.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/worksheet/range.hpp>
#include <xlnt/worksheet/range_reference.hpp>
#include <xlnt/worksheet/worksheet.hpp>
#include <xlnt/writer/worksheet_writer.hpp>

#include "detail/constants.hpp"
#include "detail/include_pugixml.hpp"

namespace {

bool is_integral(long double d)
{
    return d == static_cast<long long int>(d);
}
    
} // namepsace

namespace xlnt {

std::string write_worksheet(worksheet ws, const std::vector<std::string> &string_table)
{
    ws.get_cell("A1");
    
    pugi::xml_document doc;
    auto root_node = doc.append_child("worksheet");
    root_node.append_attribute("xmlns").set_value(constants::Namespaces.at("spreadsheetml").c_str());
    root_node.append_attribute("xmlns:r").set_value(constants::Namespaces.at("r").c_str());
    auto sheet_pr_node = root_node.append_child("sheetPr");
    auto outline_pr_node = sheet_pr_node.append_child("outlinePr");
    if(!ws.get_page_setup().is_default())
    {
        auto page_set_up_pr_node = sheet_pr_node.append_child("pageSetUpPr");
        page_set_up_pr_node.append_attribute("fitToPage").set_value(ws.get_page_setup().fit_to_page() ? 1 : 0);
    }
    outline_pr_node.append_attribute("summaryBelow").set_value(1);
    outline_pr_node.append_attribute("summaryRight").set_value(1);
    auto dimension_node = root_node.append_child("dimension");
    dimension_node.append_attribute("ref").set_value(ws.calculate_dimension().to_string().c_str());
    auto sheet_views_node = root_node.append_child("sheetViews");
    auto sheet_view_node = sheet_views_node.append_child("sheetView");
    sheet_view_node.append_attribute("workbookViewId").set_value(0);
    
    std::string active_pane = "bottomRight";
    
    if(ws.has_frozen_panes())
    {
        auto pane_node = sheet_view_node.append_child("pane");
        
        if(ws.get_frozen_panes().get_column_index() > 1)
        {
            pane_node.append_attribute("xSplit").set_value(ws.get_frozen_panes().get_column_index() - 1);
            active_pane = "topRight";
        }
        
        if(ws.get_frozen_panes().get_row() > 1)
        {
            pane_node.append_attribute("ySplit").set_value(ws.get_frozen_panes().get_row() - 1);
            active_pane = "bottomLeft";
        }
        
        if(ws.get_frozen_panes().get_row() > 1 && ws.get_frozen_panes().get_column_index() > 1)
        {
            auto top_right_node = sheet_view_node.append_child("selection");
            top_right_node.append_attribute("pane").set_value("topRight");
            auto bottom_left_node = sheet_view_node.append_child("selection");
            bottom_left_node.append_attribute("pane").set_value("bottomLeft");
            active_pane = "bottomRight";
        }
        
        pane_node.append_attribute("topLeftCell").set_value(ws.get_frozen_panes().to_string().c_str());
        pane_node.append_attribute("activePane").set_value(active_pane.c_str());
        pane_node.append_attribute("state").set_value("frozen");
    }
    
    auto selection_node = sheet_view_node.append_child("selection");
    if(ws.has_frozen_panes())
    {
        if(ws.get_frozen_panes().get_row() > 1 && ws.get_frozen_panes().get_column_index() > 1)
        {
            selection_node.append_attribute("pane").set_value("bottomRight");
        }
        else if(ws.get_frozen_panes().get_row() > 1)
        {
            selection_node.append_attribute("pane").set_value("bottomLeft");
        }
        else if(ws.get_frozen_panes().get_column_index() > 1)
        {
            selection_node.append_attribute("pane").set_value("topRight");
        }
    }
    std::string active_cell = "A1";
    selection_node.append_attribute("activeCell").set_value(active_cell.c_str());
    selection_node.append_attribute("sqref").set_value(active_cell.c_str());
    
    auto sheet_format_pr_node = root_node.append_child("sheetFormatPr");
    sheet_format_pr_node.append_attribute("baseColWidth").set_value(10);
    sheet_format_pr_node.append_attribute("defaultRowHeight").set_value(15);
    
    std::unordered_map<std::string, std::string> hyperlink_references;
    
    auto sheet_data_node = root_node.append_child("sheetData");
    
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
        
        auto row_node = sheet_data_node.append_child("row");
        row_node.append_attribute("r").set_value(row.front().get_row());
        
        row_node.append_attribute("spans").set_value((std::to_string(min) + ":" + std::to_string(max)).c_str());
        if(ws.has_row_properties(row.front().get_row()))
        {
            row_node.append_attribute("customHeight").set_value(1);
            auto height = ws.get_row_properties(row.front().get_row()).height;
            if(height == std::floor(height))
            {
                row_node.append_attribute("ht").set_value((std::to_string((int)height) + ".0").c_str());
            }
            else
            {
                row_node.append_attribute("ht").set_value(height);
            }
        }
        //row_node.append_attribute("x14ac:dyDescent").set_value(0.25);
        
        for(auto cell : row)
        {
            if(!cell.garbage_collectible())
            {
                if(cell.has_hyperlink())
                {
                    hyperlink_references[cell.get_hyperlink().get_id()] = cell.get_reference().to_string();
                }
                
                auto cell_node = row_node.append_child("c");
                cell_node.append_attribute("r").set_value(cell.get_reference().to_string().c_str());
                
                if(cell.get_data_type() == cell::type::string)
                {
                    if(cell.has_formula())
                    {
                        cell_node.append_attribute("t").set_value("str");
                        cell_node.append_child("f").text().set(cell.get_formula().c_str());
                        cell_node.append_child("v").text().set(cell.to_string().c_str());
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
                            cell_node.append_attribute("t").set_value("s");
                        }
                        else
                        {
                            cell_node.append_attribute("t").set_value("inlineStr");
                            auto inline_string_node = cell_node.append_child("is");
                            inline_string_node.append_child("t").text().set(cell.get_value<std::string>().c_str());
                        }
                    }
                    else
                    {
                        cell_node.append_attribute("t").set_value("s");
                        auto value_node = cell_node.append_child("v");
                        value_node.text().set(match_index);
                    }
                }
                else
                {
                    if(cell.get_data_type() != cell::type::null)
                    {
                        if(cell.get_data_type() == cell::type::boolean)
                        {
                            cell_node.append_attribute("t").set_value("b");
                            auto value_node = cell_node.append_child("v");
                            value_node.text().set(cell.get_value<bool>() ? 1 : 0);
                        }
                        else if(cell.get_data_type() == cell::type::numeric)
                        {
                            if(cell.has_formula())
                            {
                                cell_node.append_child("f").text().set(cell.get_formula().c_str());
                                cell_node.append_child("v").text().set(cell.to_string().c_str());
                                continue;
                            }
                            
                            cell_node.append_attribute("t").set_value("n");
                            auto value_node = cell_node.append_child("v");
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
                                value_node.text().set(ss.str().c_str());
                            }
                        }
                    }
                    else if(cell.has_formula())
                    {
                        cell_node.append_child("f").text().set(cell.get_formula().c_str());
                        cell_node.append_child("v");
                        continue;
                    }
                }
                
                //if(cell.has_style())
                {
                    auto style_id = cell.get_style_id();
                    cell_node.append_attribute("s").set_value((int)style_id);
                }
            }
        }
    }
    
    if(ws.has_auto_filter())
    {
        auto auto_filter_node = root_node.append_child("autoFilter");
        auto_filter_node.append_attribute("ref").set_value(ws.get_auto_filter().to_string().c_str());
    }
    
    if(!ws.get_merged_ranges().empty())
    {
        auto merge_cells_node = root_node.append_child("mergeCells");
        merge_cells_node.append_attribute("count").set_value((unsigned int)ws.get_merged_ranges().size());
        
        for(auto merged_range : ws.get_merged_ranges())
        {
            auto merge_cell_node = merge_cells_node.append_child("mergeCell");
            merge_cell_node.append_attribute("ref").set_value(merged_range.to_string().c_str());
        }
    }
    
    if(!ws.get_relationships().empty())
    {
        auto hyperlinks_node = root_node.append_child("hyperlinks");
        
        for(auto relationship : ws.get_relationships())
        {
            auto hyperlink_node = hyperlinks_node.append_child("hyperlink");
            hyperlink_node.append_attribute("display").set_value(relationship.get_target_uri().c_str());
            hyperlink_node.append_attribute("ref").set_value(hyperlink_references.at(relationship.get_id()).c_str());
            hyperlink_node.append_attribute("r:id").set_value(relationship.get_id().c_str());
        }
    }
    
    if(!ws.get_page_setup().is_default())
    {
        auto print_options_node = root_node.append_child("printOptions");
        print_options_node.append_attribute("horizontalCentered").set_value(ws.get_page_setup().get_horizontal_centered() ? 1 : 0);
        print_options_node.append_attribute("verticalCentered").set_value(ws.get_page_setup().get_vertical_centered() ? 1 : 0);
    }
    
    auto page_margins_node = root_node.append_child("pageMargins");
    
    page_margins_node.append_attribute("left").set_value(ws.get_page_margins().get_left());
    page_margins_node.append_attribute("right").set_value(ws.get_page_margins().get_right());
    page_margins_node.append_attribute("top").set_value(ws.get_page_margins().get_top());
    page_margins_node.append_attribute("bottom").set_value(ws.get_page_margins().get_bottom());
    page_margins_node.append_attribute("header").set_value(ws.get_page_margins().get_header());
    page_margins_node.append_attribute("footer").set_value(ws.get_page_margins().get_footer());
    
    if(!ws.get_page_setup().is_default())
    {
        auto page_setup_node = root_node.append_child("pageSetup");
        
        std::string orientation_string = ws.get_page_setup().get_orientation() == page_setup::orientation::landscape ? "landscape" : "portrait";
        page_setup_node.append_attribute("orientation").set_value(orientation_string.c_str());
        page_setup_node.append_attribute("paperSize").set_value((int)ws.get_page_setup().get_paper_size());
        page_setup_node.append_attribute("fitToHeight").set_value(ws.get_page_setup().fit_to_height() ? 1 : 0);
        page_setup_node.append_attribute("fitToWidth").set_value(ws.get_page_setup().fit_to_width() ? 1 : 0);
    }
    
    if(!ws.get_header_footer().is_default())
    {
        auto header_footer_node = root_node.append_child("headerFooter");
        auto odd_header_node = header_footer_node.append_child("oddHeader");
        std::string header_text = "&L&\"Calibri,Regular\"&K000000Left Header Text&C&\"Arial,Regular\"&6&K445566Center Header Text&R&\"Arial,Bold\"&8&K112233Right Header Text";
        odd_header_node.text().set(header_text.c_str());
        auto odd_footer_node = header_footer_node.append_child("oddFooter");
        std::string footer_text = "&L&\"Times New Roman,Regular\"&10&K445566Left Footer Text_x000D_And &D and &T&C&\"Times New Roman,Bold\"&12&K778899Center Footer Text &Z&F on &A&R&\"Times New Roman,Italic\"&14&KAABBCCRight Footer Text &P of &N";
        odd_footer_node.text().set(footer_text.c_str());
    }
    
    std::stringstream ss;
    doc.save(ss);
    
    return ss.str();
}

} // namespace xlnt
