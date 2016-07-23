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
#include <pugixml.hpp>
#include <sstream>

#include <detail/constants.hpp>
#include <detail/stylesheet.hpp>
#include <detail/worksheet_serializer.hpp>
#include <xlnt/cell/cell.hpp>
#include <xlnt/cell/cell_reference.hpp>
#include <xlnt/cell/text.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/worksheet/cell_iterator.hpp>
#include <xlnt/worksheet/column_properties.hpp>
#include <xlnt/worksheet/range.hpp>
#include <xlnt/worksheet/range_iterator.hpp>
#include <xlnt/worksheet/range_reference.hpp>
#include <xlnt/worksheet/row_properties.hpp>

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

bool worksheet_serializer::read_worksheet(const pugi::xml_document &xml, detail::stylesheet &stylesheet)
{
    auto root_node = xml.child("worksheet");

    auto dimension_node = root_node.child("dimension");
    std::string dimension(dimension_node.attribute("ref").value());
    auto full_range = xlnt::range_reference(dimension);
    auto sheet_data_node = root_node.child("sheetData");

    if (root_node.child("mergeCells"))
    {
        auto merge_cells_node = root_node.child("mergeCells");
        auto count = std::stoull(merge_cells_node.attribute("count").value());

        for (auto merge_cell_node : merge_cells_node.children("mergeCell"))
        {
            sheet_.merge_cells(range_reference(merge_cell_node.attribute("ref").value()));
            count--;
        }

        if (count != 0)
        {
            throw std::runtime_error("mismatch between count and actual number of merged cells");
        }
    }

    auto &shared_strings = sheet_.get_workbook().get_shared_strings();

    for (auto row_node : sheet_data_node.children("row"))
    {
        auto row_index = static_cast<row_t>(std::stoull(row_node.attribute("r").value()));

        if (row_node.attribute("ht"))
        {
            sheet_.get_row_properties(row_index).height = std::stold(row_node.attribute("ht").value());
        }

        std::string span_string = row_node.attribute("spans").value();
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

            pugi::xml_node cell_node;
            bool cell_found = false;

            for (auto &check_cell_node : row_node.children("c"))
            {
                if (check_cell_node.attribute("r") && check_cell_node.attribute("r").value() == address)
                {
                    cell_node = check_cell_node;
                    cell_found = true;
                    break;
                }
            }

            if (cell_found)
            {
                bool has_value = cell_node.child("v") != nullptr;
                std::string value_string = has_value ? cell_node.child("v").text().get() : "";

                bool has_type = cell_node.attribute("t") != nullptr;
                std::string type = has_type ? cell_node.attribute("t").value() : "";

                bool has_format = cell_node.attribute("s") != nullptr;
                auto format_id = static_cast<std::size_t>(has_format ? std::stoull(cell_node.attribute("s").value()) : 0LL);

                bool has_formula = cell_node.child("f") != nullptr;
                bool has_shared_formula = has_formula && cell_node.child("f").attribute("t") != nullptr
                    && cell_node.child("f").attribute("t").value() == std::string("shared");

                auto cell = sheet_.get_cell(address);

                if (has_formula && !has_shared_formula && !sheet_.get_workbook().get_data_only())
                {
                    std::string formula = cell_node.child("f").text().get();
                    cell.set_formula(formula);
                }

                if (has_type && type == "inlineStr") // inline string
                {
                    std::string inline_string = cell_node.child("is").child("t").text().get();
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

                if (has_format)
                {
                    cell.set_format(stylesheet.formats.at(format_id));
                }
            }
        }
    }

    auto cols_node = root_node.child("cols");

    for (auto col_node : cols_node.children("col"))
    {
        auto min = static_cast<column_t::index_t>(std::stoull(col_node.attribute("min").value()));
        auto max = static_cast<column_t::index_t>(std::stoull(col_node.attribute("max").value()));
        auto width = std::stold(col_node.attribute("width").value());
        bool custom = col_node.attribute("customWidth").value() == std::string("1");
        auto column_style = static_cast<std::size_t>(col_node.attribute("style") ? std::stoull(col_node.attribute("style").value()) : 0);

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

    if (root_node.child("autoFilter"))
    {
        auto auto_filter_node = root_node.child("autoFilter");
        xlnt::range_reference ref(auto_filter_node.attribute("ref").value());
        sheet_.auto_filter(ref);
    }

    return true;
}

void worksheet_serializer::write_worksheet(pugi::xml_document &xml) const
{
    auto root_node = xml.append_child("worksheet");

    root_node.append_attribute("xmlns").set_value(constants::get_namespace("spreadsheetml").c_str());
    root_node.append_attribute("xmlns:r").set_value(constants::get_namespace("r").c_str());

    auto sheet_pr_node = root_node.append_child("sheetPr");

    auto outline_pr_node = sheet_pr_node.append_child("outlinePr");

    outline_pr_node.append_attribute("summaryBelow").set_value("1");
    outline_pr_node.append_attribute("summaryRight").set_value("1");
    
    if (!sheet_.get_page_setup().is_default())
    {
        auto page_set_up_pr_node = sheet_pr_node.append_child("pageSetUpPr");
        page_set_up_pr_node.append_attribute("fitToPage").set_value(sheet_.get_page_setup().fit_to_page() ? "1" : "0");
    }

    auto dimension_node = root_node.append_child("dimension");
    dimension_node.append_attribute("ref").set_value(sheet_.calculate_dimension().to_string().c_str());

    auto sheet_views_node = root_node.append_child("sheetViews");
    auto sheet_view_node = sheet_views_node.append_child("sheetView");
    sheet_view_node.append_attribute("workbookViewId").set_value("0");

    std::string active_pane = "bottomRight";

    if (sheet_.has_frozen_panes())
    {
        auto pane_node = sheet_view_node.append_child("pane");

        if (sheet_.get_frozen_panes().get_column_index() > 1)
        {
            pane_node.append_attribute("xSplit").set_value(std::to_string(sheet_.get_frozen_panes().get_column_index().index - 1).c_str());
            active_pane = "topRight";
        }

        if (sheet_.get_frozen_panes().get_row() > 1)
        {
            pane_node.append_attribute("ySplit").set_value(std::to_string(sheet_.get_frozen_panes().get_row() - 1).c_str());
            active_pane = "bottomLeft";
        }

        if (sheet_.get_frozen_panes().get_row() > 1 && sheet_.get_frozen_panes().get_column_index() > 1)
        {
            auto top_right_node = sheet_view_node.append_child("selection");
            top_right_node.append_attribute("pane").set_value("topRight");
            auto bottom_left_node = sheet_view_node.append_child("selection");
            bottom_left_node.append_attribute("pane").set_value("bottomLeft");
            active_pane = "bottomRight";
        }

        pane_node.append_attribute("topLeftCell").set_value(sheet_.get_frozen_panes().to_string().c_str());
        pane_node.append_attribute("activePane").set_value(active_pane.c_str());
        pane_node.append_attribute("state").set_value("frozen");
    }

    auto selection_node = sheet_view_node.append_child("selection");

    if (sheet_.has_frozen_panes())
    {
        if (sheet_.get_frozen_panes().get_row() > 1 && sheet_.get_frozen_panes().get_column_index() > 1)
        {
            selection_node.append_attribute("pane").set_value("bottomRight");
        }
        else if (sheet_.get_frozen_panes().get_row() > 1)
        {
            selection_node.append_attribute("pane").set_value("bottomLeft");
        }
        else if (sheet_.get_frozen_panes().get_column_index() > 1)
        {
            selection_node.append_attribute("pane").set_value("topRight");
        }
    }

    std::string active_cell = "A1";
    selection_node.append_attribute("activeCell").set_value(active_cell.c_str());
    selection_node.append_attribute("sqref").set_value(active_cell.c_str());

    auto sheet_format_pr_node = root_node.append_child("sheetFormatPr");
    sheet_format_pr_node.append_attribute("baseColWidth").set_value("10");
    sheet_format_pr_node.append_attribute("defaultRowHeight").set_value("15");

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
        auto cols_node = root_node.append_child("cols");

        for (auto column = sheet_.get_lowest_column(); column <= sheet_.get_highest_column(); column++)
        {
            if(!sheet_.has_column_properties(column))
            {
                continue;
            }
            
            const auto &props = sheet_.get_column_properties(column);

            auto col_node = cols_node.append_child("col");

            col_node.append_attribute("min").set_value(std::to_string(column.index).c_str());
            col_node.append_attribute("max").set_value(std::to_string(column.index).c_str());
            col_node.append_attribute("width").set_value(std::to_string(props.width).c_str());
            col_node.append_attribute("style").set_value(std::to_string(props.style).c_str());
            col_node.append_attribute("customWidth").set_value(props.custom ? "1" : "0");
        }
    }

    std::unordered_map<std::string, std::string> hyperlink_references;

    auto sheet_data_node = root_node.append_child("sheetData");
    const auto &shared_strings = sheet_.get_workbook().get_shared_strings();

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

        auto row_node = sheet_data_node.append_child("row");

        row_node.append_attribute("r").set_value(std::to_string(row.front().get_row()).c_str());
        row_node.append_attribute("spans").set_value((std::to_string(min) + ":" + std::to_string(max)).c_str());

        if (sheet_.has_row_properties(row.front().get_row()))
        {
            row_node.append_attribute("customHeight").set_value("1");
            auto height = sheet_.get_row_properties(row.front().get_row()).height;

            if (height == std::floor(height))
            {
                row_node.append_attribute("ht").set_value((std::to_string(static_cast<long long int>(height)) + ".0").c_str());
            }
            else
            {
                row_node.append_attribute("ht").set_value(std::to_string(height).c_str());
            }
        }

        // row_node.append_attribute("x14ac:dyDescent", 0.25);

        for (auto cell : row)
        {
            if (!cell.garbage_collectible())
            {
                if (cell.has_hyperlink())
                {
                    hyperlink_references[cell.get_hyperlink().get_id()] = cell.get_reference().to_string();
                }

                auto cell_node = row_node.append_child("c");
                cell_node.append_attribute("r").set_value(cell.get_reference().to_string().c_str());

                if (cell.get_data_type() == cell::type::string)
                {
                    if (cell.has_formula())
                    {
                        cell_node.append_attribute("t").set_value("str");
                        cell_node.append_child("f").text().set(cell.get_formula().c_str());
                        cell_node.append_child("v").text().set(cell.to_string().c_str());

                        continue;
                    }

                    int match_index = -1;

                    for (std::size_t i = 0; i < shared_strings.size(); i++)
                    {
                        if (shared_strings[i] == cell.get_value<text>())
                        {
                            match_index = static_cast<int>(i);
                            break;
                        }
                    }

                    if (match_index == -1)
                    {
                        if (cell.get_value<std::string>().empty())
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
                        value_node.text().set(std::to_string(match_index).c_str());
                    }
                }
                else
                {
                    if (cell.get_data_type() != cell::type::null)
                    {
                        if (cell.get_data_type() == cell::type::boolean)
                        {
                            cell_node.append_attribute("t").set_value("b");
                            auto value_node = cell_node.append_child("v");
                            value_node.text().set(cell.get_value<bool>() ? "1" : "0");
                        }
                        else if (cell.get_data_type() == cell::type::numeric)
                        {
                            if (cell.has_formula())
                            {
                                cell_node.append_child("f").text().set(cell.get_formula().c_str());
                                cell_node.append_child("v").text().set(cell.to_string().c_str());
                                continue;
                            }

                            cell_node.append_attribute("t").set_value("n");
                            auto value_node = cell_node.append_child("v");

                            if (is_integral(cell.get_value<long double>()))
                            {
                                value_node.text().set(std::to_string(cell.get_value<long long>()).c_str());
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
                    else if (cell.has_formula())
                    {
                        cell_node.append_child("f").text().set(cell.get_formula().c_str());
                        cell_node.append_child("v");
                        continue;
                    }
                }

                if (cell.has_format())
                {
                    cell_node.append_attribute("s").set_value(std::to_string(cell.get_format_id()).c_str());
                }
            }
        }
    }

    if (sheet_.has_auto_filter())
    {
        auto auto_filter_node = root_node.append_child("autoFilter");
        auto_filter_node.append_attribute("ref").set_value(sheet_.get_auto_filter().to_string().c_str());
    }

    if (!sheet_.get_merged_ranges().empty())
    {
        auto merge_cells_node = root_node.append_child("mergeCells");
        merge_cells_node.append_attribute("count").set_value(std::to_string(sheet_.get_merged_ranges().size()).c_str());

        for (auto merged_range : sheet_.get_merged_ranges())
        {
            auto merge_cell_node = merge_cells_node.append_child("mergeCell");
            merge_cell_node.append_attribute("ref").set_value(merged_range.to_string().c_str());
        }
    }

    if (!sheet_.get_relationships().empty())
    {
        auto hyperlinks_node = root_node.append_child("hyperlinks");

        for (const auto &relationship : sheet_.get_relationships())
        {
            auto hyperlink_node = hyperlinks_node.append_child("hyperlink");
            hyperlink_node.append_attribute("display").set_value(relationship.get_target_uri().c_str());
            hyperlink_node.append_attribute("ref").set_value(hyperlink_references.at(relationship.get_id()).c_str());
            hyperlink_node.append_attribute("r:id").set_value(relationship.get_id().c_str());
        }
    }

    if (!sheet_.get_page_setup().is_default())
    {
        auto print_options_node = root_node.append_child("printOptions");
        print_options_node.append_attribute("horizontalCentered").set_value(
                                         sheet_.get_page_setup().get_horizontal_centered() ? "1" : "0");
        print_options_node.append_attribute("verticalCentered").set_value(
                                         sheet_.get_page_setup().get_vertical_centered() ? "1" : "0");
    }

    auto page_margins_node = root_node.append_child("pageMargins");
    
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

    page_margins_node.append_attribute("left").set_value(remove_trailing_zeros(std::to_string(sheet_.get_page_margins().get_left())).c_str());
    page_margins_node.append_attribute("right").set_value(remove_trailing_zeros(std::to_string(sheet_.get_page_margins().get_right())).c_str());
    page_margins_node.append_attribute("top").set_value(remove_trailing_zeros(std::to_string(sheet_.get_page_margins().get_top())).c_str());
    page_margins_node.append_attribute("bottom").set_value(remove_trailing_zeros(std::to_string(sheet_.get_page_margins().get_bottom())).c_str());
    page_margins_node.append_attribute("header").set_value(remove_trailing_zeros(std::to_string(sheet_.get_page_margins().get_header())).c_str());
    page_margins_node.append_attribute("footer").set_value(remove_trailing_zeros(std::to_string(sheet_.get_page_margins().get_footer())).c_str());

    if (!sheet_.get_page_setup().is_default())
    {
        auto page_setup_node = root_node.append_child("pageSetup");

        std::string orientation_string =
            sheet_.get_page_setup().get_orientation() == orientation::landscape ? "landscape" : "portrait";
        page_setup_node.append_attribute("orientation").set_value(orientation_string.c_str());
        page_setup_node.append_attribute("paperSize").set_value(
                                      std::to_string(static_cast<int>(sheet_.get_page_setup().get_paper_size())).c_str());
        page_setup_node.append_attribute("fitToHeight").set_value(sheet_.get_page_setup().fit_to_height() ? "1" : "0");
        page_setup_node.append_attribute("fitToWidth").set_value(sheet_.get_page_setup().fit_to_width() ? "1" : "0");
    }

    if (!sheet_.get_header_footer().is_default())
    {
        auto header_footer_node = root_node.append_child("headerFooter");
        auto odd_header_node = header_footer_node.append_child("oddHeader");
        std::string header_text =
            "&L&\"Calibri,Regular\"&K000000Left Header Text&C&\"Arial,Regular\"&6&K445566Center Header "
            "Text&R&\"Arial,Bold\"&8&K112233Right Header Text";
        odd_header_node.text().set(header_text.c_str());
        auto odd_footer_node = header_footer_node.append_child("oddFooter");
        std::string footer_text =
            "&L&\"Times New Roman,Regular\"&10&K445566Left Footer Text_x000D_And &D and &T&C&\"Times New "
            "Roman,Bold\"&12&K778899Center Footer Text &Z&F on &A&R&\"Times New Roman,Italic\"&14&KAABBCCRight Footer "
            "Text &P of &N";
        odd_footer_node.text().set(footer_text.c_str());
    }
}

} // namespace xlnt
