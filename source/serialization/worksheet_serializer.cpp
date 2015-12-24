// Copyright (c) 2014-2016 Thomas Fussell
// Copyright (c) 2010-2015 openpyxl
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, WRISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE
//
// @license: http://www.opensource.org/licenses/mit-license.php
// @author: see AUTHORS file
#include <algorithm>
#include <cmath>
#include <sstream>

#include <xlnt/serialization/worksheet_serializer.hpp>
#include <xlnt/cell/cell.hpp>
#include <xlnt/cell/cell_reference.hpp>
#include <xlnt/serialization/xml_document.hpp>
#include <xlnt/serialization/xml_node.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/worksheet/column_properties.hpp>
#include <xlnt/worksheet/range.hpp>
#include <xlnt/worksheet/range_reference.hpp>
#include <xlnt/worksheet/row_properties.hpp>

#include <detail/constants.hpp>

namespace {

bool is_integral(long double d)
{
    return d == static_cast<long long int>(d);
}

} // namepsace

namespace xlnt {

worksheet_serializer::worksheet_serializer(worksheet sheet) : sheet_(sheet)
{
}

bool worksheet_serializer::read_worksheet(const xml_document &xml)
{
    auto &root_node = xml.get_child("worksheet");

    auto &dimension_node = root_node.get_child("dimension");
    std::string dimension = dimension_node.get_attribute("ref");
    auto full_range = xlnt::range_reference(dimension);
    auto sheet_data_node = root_node.get_child("sheetData");

    if (root_node.has_child("mergeCells"))
    {
        auto merge_cells_node = root_node.get_child("mergeCells");
        auto count = std::stoull(merge_cells_node.get_attribute("count"));

        for (auto merge_cell_node : merge_cells_node.get_children())
        {
            if (merge_cell_node.get_name() != "mergeCell")
            {
                continue;
            }

            sheet_.merge_cells(range_reference(merge_cell_node.get_attribute("ref")));
            count--;
        }

        if (count != 0)
        {
            throw std::runtime_error("mismatch between count and actual number of merged cells");
        }
    }

    auto &shared_strings = sheet_.get_parent().get_shared_strings();

    for (auto row_node : sheet_data_node.get_children())
    {
        if (row_node.get_name() != "row")
        {
            continue;
        }

        auto row_index = static_cast<row_t>(std::stoull(row_node.get_attribute("r")));

        if (row_node.has_attribute("ht"))
        {
            sheet_.get_row_properties(row_index).height = std::stold(row_node.get_attribute("ht"));
        }

        std::string span_string = row_node.get_attribute("spans");
        auto colon_index = span_string.find(':');

        column_t min_column = 0;
        column_t max_column = 0;

        if (colon_index != std::string::npos)
        {
            min_column = static_cast<column_t::index_t>(std::stoll(span_string.substr(0, colon_index)));
            max_column = static_cast<column_t::index_t>(std::stoll(span_string.substr(colon_index + 1)));
        }
        else
        {
            min_column = full_range.get_top_left().get_column_index();
            max_column = full_range.get_bottom_right().get_column_index();
        }

        for (column_t i = min_column; i <= max_column; i++)
        {
            std::string address = i.column_string() + std::to_string(row_index);

            xml_node cell_node;
            bool cell_found = false;

            for (auto &check_cell_node : row_node.get_children())
            {
                if (check_cell_node.get_name() != "c")
                {
                    continue;
                }

                if (check_cell_node.has_attribute("r") && check_cell_node.get_attribute("r") == address)
                {
                    cell_node = check_cell_node;
                    cell_found = true;
                    break;
                }
            }

            if (cell_found)
            {
                bool has_value = cell_node.has_child("v");
                std::string value_string = has_value ? cell_node.get_child("v").get_text() : "";

                bool has_type = cell_node.has_attribute("t");
                std::string type = has_type ? cell_node.get_attribute("t") : "";

                bool has_style = cell_node.has_attribute("s");
                auto style_id = static_cast<std::size_t>(has_style ? std::stoull(cell_node.get_attribute("s")) : 0LL);

                bool has_formula = cell_node.has_child("f");
                bool has_shared_formula = has_formula && cell_node.get_child("f").has_attribute("t") &&
                                          cell_node.get_child("f").get_attribute("t") == "shared";

                auto cell = sheet_.get_cell(address);

                if (has_formula && !has_shared_formula && !sheet_.get_parent().get_data_only())
                {
                    std::string formula = cell_node.get_child("f").get_text();
                    cell.set_formula(formula);
                }

                if (has_type && type == "inlineStr") // inline string
                {
                    std::string inline_string = cell_node.get_child("is").get_child("t").get_text();
                    cell.set_value(inline_string);
                }
                else if (has_type && type == "s" && !has_formula) // shared string
                {
                    auto shared_string_index = static_cast<std::size_t>(std::stoull(value_string));
                    auto shared_string = shared_strings.at(shared_string_index);
                    cell.set_value(shared_string);
                }
                else if (has_type && type == "b") // boolean
                {
                    cell.set_value(value_string != "0");
                }
                else if (has_type && type == "str")
                {
                    cell.set_value(value_string);
                }
                else if (has_value && !value_string.empty())
                {
                    if (!value_string.empty() && value_string[0] == '#')
                    {
                        cell.set_error(value_string);
                    }
                    else
                    {
                        cell.set_value(std::stold(value_string));
                    }
                }

                if (has_style)
                {
                    cell.set_style_id(style_id);
                }
            }
        }
    }

    auto &cols_node = root_node.get_child("cols");

    for (auto col_node : cols_node.get_children())
    {
        if (col_node.get_name() != "col")
        {
            continue;
        }

        auto min = static_cast<column_t::index_t>(std::stoull(col_node.get_attribute("min")));
        auto max = static_cast<column_t::index_t>(std::stoull(col_node.get_attribute("max")));
        auto width = std::stold(col_node.get_attribute("width"));
        bool custom = col_node.get_attribute("customWidth") == "1";
        auto column_style = static_cast<std::size_t>(col_node.has_attribute("style") ? std::stoull(col_node.get_attribute("style")) : 0);

        for (auto column = min; column <= max; column++)
        {
            if (!sheet_.has_column_properties(column))
            {
                sheet_.add_column_properties(column, column_properties());
            }

            sheet_.get_column_properties(min).width = width;
            sheet_.get_column_properties(min).style = column_style;
            sheet_.get_column_properties(min).custom = custom;
        }
    }

    if (root_node.has_child("autoFilter"))
    {
        auto &auto_filter_node = root_node.get_child("autoFilter");
        xlnt::range_reference ref(auto_filter_node.get_attribute("ref"));
        sheet_.auto_filter(ref);
    }

    return true;
}

xml_document worksheet_serializer::write_worksheet() const
{
    xml_document xml;

    auto root_node = xml.add_child("worksheet");

    xml.add_namespace("", constants::Namespace("spreadsheetml"));
    xml.add_namespace("r", constants::Namespace("r"));

    auto sheet_pr_node = root_node.add_child("sheetPr");

    auto outline_pr_node = sheet_pr_node.add_child("outlinePr");

    outline_pr_node.add_attribute("summaryBelow", "1");
    outline_pr_node.add_attribute("summaryRight", "1");
    
    if (!sheet_.get_page_setup().is_default())
    {
        auto page_set_up_pr_node = sheet_pr_node.add_child("pageSetUpPr");
        page_set_up_pr_node.add_attribute("fitToPage", sheet_.get_page_setup().fit_to_page() ? "1" : "0");
    }

    auto dimension_node = root_node.add_child("dimension");
    dimension_node.add_attribute("ref", sheet_.calculate_dimension().to_string());

    auto sheet_views_node = root_node.add_child("sheetViews");
    auto sheet_view_node = sheet_views_node.add_child("sheetView");
    sheet_view_node.add_attribute("workbookViewId", "0");

    std::string active_pane = "bottomRight";

    if (sheet_.has_frozen_panes())
    {
        auto pane_node = sheet_view_node.add_child("pane");

        if (sheet_.get_frozen_panes().get_column_index() > 1)
        {
            pane_node.add_attribute("xSplit", std::to_string(sheet_.get_frozen_panes().get_column_index().index - 1));
            active_pane = "topRight";
        }

        if (sheet_.get_frozen_panes().get_row() > 1)
        {
            pane_node.add_attribute("ySplit", std::to_string(sheet_.get_frozen_panes().get_row() - 1));
            active_pane = "bottomLeft";
        }

        if (sheet_.get_frozen_panes().get_row() > 1 && sheet_.get_frozen_panes().get_column_index() > 1)
        {
            auto top_right_node = sheet_view_node.add_child("selection");
            top_right_node.add_attribute("pane", "topRight");
            auto bottom_left_node = sheet_view_node.add_child("selection");
            bottom_left_node.add_attribute("pane", "bottomLeft");
            active_pane = "bottomRight";
        }

        pane_node.add_attribute("topLeftCell", sheet_.get_frozen_panes().to_string());
        pane_node.add_attribute("activePane", active_pane);
        pane_node.add_attribute("state", "frozen");
    }

    auto selection_node = sheet_view_node.add_child("selection");

    if (sheet_.has_frozen_panes())
    {
        if (sheet_.get_frozen_panes().get_row() > 1 && sheet_.get_frozen_panes().get_column_index() > 1)
        {
            selection_node.add_attribute("pane", "bottomRight");
        }
        else if (sheet_.get_frozen_panes().get_row() > 1)
        {
            selection_node.add_attribute("pane", "bottomLeft");
        }
        else if (sheet_.get_frozen_panes().get_column_index() > 1)
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

    for (auto column = sheet_.get_lowest_column(); column <= sheet_.get_highest_column(); column++)
    {
        if (sheet_.has_column_properties(column))
        {
            has_column_properties = true;
            break;
        }
    }

    if (has_column_properties)
    {
        auto cols_node = root_node.add_child("cols");

        for (auto column = sheet_.get_lowest_column(); column <= sheet_.get_highest_column(); column++)
        {
            if(!sheet_.has_column_properties(column))
            {
                continue;
            }
            
            const auto &props = sheet_.get_column_properties(column);

            auto col_node = cols_node.add_child("col");

            col_node.add_attribute("min", std::to_string(column.index));
            col_node.add_attribute("max", std::to_string(column.index));
            col_node.add_attribute("width", std::to_string(props.width));
            col_node.add_attribute("style", std::to_string(props.style));
            col_node.add_attribute("customWidth", props.custom ? "1" : "0");
        }
    }

    std::unordered_map<std::string, std::string> hyperlink_references;

    auto sheet_data_node = root_node.add_child("sheetData");
    const auto &shared_strings = sheet_.get_parent().get_shared_strings();

    for (auto row : sheet_.rows())
    {
        row_t min = static_cast<row_t>(row.num_cells());
        row_t max = 0;
        bool any_non_null = false;

        for (auto cell : row)
        {
            min = std::min(min, cell.get_column().index);
            max = std::max(max, cell.get_column().index);

            if (!cell.garbage_collectible())
            {
                any_non_null = true;
            }
        }

        if (!any_non_null)
        {
            continue;
        }

        auto row_node = sheet_data_node.add_child("row");

        row_node.add_attribute("r", std::to_string(row.front().get_row()));
        row_node.add_attribute("spans", (std::to_string(min) + ":" + std::to_string(max)));

        if (sheet_.has_row_properties(row.front().get_row()))
        {
            row_node.add_attribute("customHeight", "1");
            auto height = sheet_.get_row_properties(row.front().get_row()).height;

            if (height == std::floor(height))
            {
                row_node.add_attribute("ht", std::to_string(static_cast<long long int>(height)) + ".0");
            }
            else
            {
                row_node.add_attribute("ht", std::to_string(height));
            }
        }

        // row_node.add_attribute("x14ac:dyDescent", 0.25);

        for (auto cell : row)
        {
            if (!cell.garbage_collectible())
            {
                if (cell.has_hyperlink())
                {
                    hyperlink_references[cell.get_hyperlink().get_id()] = cell.get_reference().to_string();
                }

                auto cell_node = row_node.add_child("c");
                cell_node.add_attribute("r", cell.get_reference().to_string());

                if (cell.get_data_type() == cell::type::string)
                {
                    if (cell.has_formula())
                    {
                        cell_node.add_attribute("t", "str");
                        cell_node.add_child("f").set_text(cell.get_formula());
                        cell_node.add_child("v").set_text(cell.to_string());

                        continue;
                    }

                    int match_index = -1;

                    for (std::size_t i = 0; i < shared_strings.size(); i++)
                    {
                        if (shared_strings[i] == cell.get_value<std::string>())
                        {
                            match_index = static_cast<int>(i);
                            break;
                        }
                    }

                    if (match_index == -1)
                    {
                        if (cell.get_value<std::string>().empty())
                        {
                            cell_node.add_attribute("t", "s");
                        }
                        else
                        {
                            cell_node.add_attribute("t", "inlineStr");
                            auto inline_string_node = cell_node.add_child("is");
                            inline_string_node.add_child("t").set_text(cell.get_value<std::string>());
                        }
                    }
                    else
                    {
                        cell_node.add_attribute("t", "s");
                        auto value_node = cell_node.add_child("v");
                        value_node.set_text(std::to_string(match_index));
                    }
                }
                else
                {
                    if (cell.get_data_type() != cell::type::null)
                    {
                        if (cell.get_data_type() == cell::type::boolean)
                        {
                            cell_node.add_attribute("t", "b");
                            auto value_node = cell_node.add_child("v");
                            value_node.set_text(cell.get_value<bool>() ? "1" : "0");
                        }
                        else if (cell.get_data_type() == cell::type::numeric)
                        {
                            if (cell.has_formula())
                            {
                                cell_node.add_child("f").set_text(cell.get_formula());
                                cell_node.add_child("v").set_text(cell.to_string());
                                continue;
                            }

                            cell_node.add_attribute("t", "n");
                            auto value_node = cell_node.add_child("v");

                            if (is_integral(cell.get_value<long double>()))
                            {
                                value_node.set_text(std::to_string(cell.get_value<long long>()));
                            }
                            else
                            {
                                std::stringstream ss;
                                ss.precision(20);
                                ss << cell.get_value<long double>();
                                ss.str();
                                value_node.set_text(ss.str());
                            }
                        }
                    }
                    else if (cell.has_formula())
                    {
                        cell_node.add_child("f").set_text(cell.get_formula());
                        cell_node.add_child("v");
                        continue;
                    }
                }

                if (cell.has_style())
                {
                    cell_node.add_attribute("s", std::to_string(cell.get_style_id()));
                }
            }
        }
    }

    if (sheet_.has_auto_filter())
    {
        auto auto_filter_node = root_node.add_child("autoFilter");
        auto_filter_node.add_attribute("ref", sheet_.get_auto_filter().to_string());
    }

    if (!sheet_.get_merged_ranges().empty())
    {
        auto merge_cells_node = root_node.add_child("mergeCells");
        merge_cells_node.add_attribute("count", std::to_string(sheet_.get_merged_ranges().size()));

        for (auto merged_range : sheet_.get_merged_ranges())
        {
            auto merge_cell_node = merge_cells_node.add_child("mergeCell");
            merge_cell_node.add_attribute("ref", merged_range.to_string());
        }
    }

    if (!sheet_.get_relationships().empty())
    {
        auto hyperlinks_node = root_node.add_child("hyperlinks");

        for (const auto &relationship : sheet_.get_relationships())
        {
            auto hyperlink_node = hyperlinks_node.add_child("hyperlink");
            hyperlink_node.add_attribute("display", relationship.get_target_uri());
            hyperlink_node.add_attribute("ref", hyperlink_references.at(relationship.get_id()));
            hyperlink_node.add_attribute("r:id", relationship.get_id());
        }
    }

    if (!sheet_.get_page_setup().is_default())
    {
        auto print_options_node = root_node.add_child("printOptions");
        print_options_node.add_attribute("horizontalCentered",
                                         sheet_.get_page_setup().get_horizontal_centered() ? "1" : "0");
        print_options_node.add_attribute("verticalCentered",
                                         sheet_.get_page_setup().get_vertical_centered() ? "1" : "0");
    }

    auto page_margins_node = root_node.add_child("pageMargins");
    
    //TODO: there must be a better way to do this
    auto remove_trailing_zeros = [](const std::string &n)
    {
        auto decimal = n.find('.');
        
        if (decimal == std::string::npos) return n;
        
        auto index = n.size() - 1;
        
        while (index >= decimal && n[index] == '0')
        {
            index--;
        }
        
        if(index == decimal)
        {
            return n.substr(0, decimal);
        }
        
        return n.substr(0, index + 1);
    };

    page_margins_node.add_attribute("left", remove_trailing_zeros(std::to_string(sheet_.get_page_margins().get_left())));
    page_margins_node.add_attribute("right", remove_trailing_zeros(std::to_string(sheet_.get_page_margins().get_right())));
    page_margins_node.add_attribute("top", remove_trailing_zeros(std::to_string(sheet_.get_page_margins().get_top())));
    page_margins_node.add_attribute("bottom", remove_trailing_zeros(std::to_string(sheet_.get_page_margins().get_bottom())));
    page_margins_node.add_attribute("header", remove_trailing_zeros(std::to_string(sheet_.get_page_margins().get_header())));
    page_margins_node.add_attribute("footer", remove_trailing_zeros(std::to_string(sheet_.get_page_margins().get_footer())));

    if (!sheet_.get_page_setup().is_default())
    {
        auto page_setup_node = root_node.add_child("pageSetup");

        std::string orientation_string =
            sheet_.get_page_setup().get_orientation() == orientation::landscape ? "landscape" : "portrait";
        page_setup_node.add_attribute("orientation", orientation_string);
        page_setup_node.add_attribute("paperSize",
                                      std::to_string(static_cast<int>(sheet_.get_page_setup().get_paper_size())));
        page_setup_node.add_attribute("fitToHeight", sheet_.get_page_setup().fit_to_height() ? "1" : "0");
        page_setup_node.add_attribute("fitToWidth", sheet_.get_page_setup().fit_to_width() ? "1" : "0");
    }

    if (!sheet_.get_header_footer().is_default())
    {
        auto header_footer_node = root_node.add_child("headerFooter");
        auto odd_header_node = header_footer_node.add_child("oddHeader");
        std::string header_text =
            "&L&\"Calibri,Regular\"&K000000Left Header Text&C&\"Arial,Regular\"&6&K445566Center Header "
            "Text&R&\"Arial,Bold\"&8&K112233Right Header Text";
        odd_header_node.set_text(header_text);
        auto odd_footer_node = header_footer_node.add_child("oddFooter");
        std::string footer_text =
            "&L&\"Times New Roman,Regular\"&10&K445566Left Footer Text_x000D_And &D and &T&C&\"Times New "
            "Roman,Bold\"&12&K778899Center Footer Text &Z&F on &A&R&\"Times New Roman,Italic\"&14&KAABBCCRight Footer "
            "Text &P of &N";
        odd_footer_node.set_text(footer_text);
    }

    return xml;
}

} // namespace xlnt
