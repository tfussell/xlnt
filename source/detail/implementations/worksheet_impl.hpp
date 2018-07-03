// Copyright (c) 2014-2018 Thomas Fussell
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

#pragma once

#include <string>
#include <unordered_map>
#include <vector>

#include <xlnt/packaging/ext_list.hpp>
#include <xlnt/workbook/named_range.hpp>
#include <xlnt/worksheet/column_properties.hpp>
#include <xlnt/worksheet/header_footer.hpp>
#include <xlnt/worksheet/phonetic_pr.hpp>
#include <xlnt/worksheet/range.hpp>
#include <xlnt/worksheet/range_reference.hpp>
#include <xlnt/worksheet/row_properties.hpp>
#include <xlnt/worksheet/sheet_format_properties.hpp>
#include <xlnt/worksheet/sheet_view.hpp>
#include <xlnt/worksheet/print_options.hpp>
#include <xlnt/worksheet/sheet_pr.hpp>
#include <detail/implementations/cell_impl.hpp>

namespace xlnt {

class workbook;

namespace detail {

struct worksheet_impl
{
    worksheet_impl(workbook *parent_workbook, std::size_t id, const std::string &title)
        : parent_(parent_workbook),
          id_(id),
          title_(title)
    {
    }

    worksheet_impl(const worksheet_impl &other)
    {
        *this = other;
    }

    void operator=(const worksheet_impl &other)
    {
        parent_ = other.parent_;

        id_ = other.id_;
        title_ = other.title_;
        format_properties_ = other.format_properties_;
        column_properties_ = other.column_properties_;
        row_properties_ = other.row_properties_;
        cell_map_ = other.cell_map_;
        page_setup_ = other.page_setup_;
        auto_filter_ = other.auto_filter_;
        page_margins_ = other.page_margins_;
        merged_cells_ = other.merged_cells_;
        named_ranges_ = other.named_ranges_;
        phonetic_properties_ = other.phonetic_properties_;
        header_footer_ = other.header_footer_;
        print_title_cols_ = other.print_title_cols_;
        print_title_rows_ = other.print_title_rows_;
        print_area_ = other.print_area_;
        views_ = other.views_;
        column_breaks_ = other.column_breaks_;
        row_breaks_ = other.row_breaks_;
        extension_list_ = other.extension_list_;
        sheet_properties_ = other.sheet_properties_;
        print_options_ = other.print_options_;

        for (auto &row : cell_map_)
        {
            for (auto &cell : row.second)
            {
                cell.second.parent_ = this;
            }
        }
    }

    workbook *parent_;

    bool operator==(const worksheet_impl& rhs) const
    {
        return id_ == rhs.id_
            && title_ == rhs.title_
            && format_properties_ == rhs.format_properties_
            && column_properties_ == rhs.column_properties_
            && row_properties_ == rhs.row_properties_
            && cell_map_ == rhs.cell_map_
            && page_setup_ == rhs.page_setup_
            && auto_filter_ == rhs.auto_filter_
            && page_margins_ == rhs.page_margins_
            && merged_cells_ == rhs.merged_cells_
            && named_ranges_ == rhs.named_ranges_
            && phonetic_properties_ == rhs.phonetic_properties_
            && header_footer_ == rhs.header_footer_
            && print_title_cols_ == rhs.print_title_cols_
            && print_title_rows_ == rhs.print_title_rows_
            && print_area_ == rhs.print_area_
            && views_ == rhs.views_
            && column_breaks_ == rhs.column_breaks_
            && row_breaks_ == rhs.row_breaks_
            && comments_ == rhs.comments_
            && print_options_ == rhs.print_options_
            && sheet_properties_ == rhs.sheet_properties_
            && extension_list_ == rhs.extension_list_;
    }

    std::size_t id_;
    std::string title_;

    sheet_format_properties format_properties_;

    std::unordered_map<column_t, column_properties> column_properties_;
    std::unordered_map<row_t, row_properties> row_properties_;

    std::unordered_map<row_t, std::unordered_map<column_t, cell_impl>> cell_map_;

    optional<page_setup> page_setup_;
    optional<range_reference> auto_filter_;
    optional<page_margins> page_margins_;
    std::vector<range_reference> merged_cells_;
    std::unordered_map<std::string, named_range> named_ranges_;

    optional<phonetic_pr> phonetic_properties_;
    optional<header_footer> header_footer_;

    std::string print_title_cols_;
    std::string print_title_rows_;

    optional<range_reference> print_area_;

    std::vector<sheet_view> views_;

    std::vector<column_t> column_breaks_;
    std::vector<row_t> row_breaks_;

    std::unordered_map<std::string, comment> comments_;
    optional<print_options> print_options_;
    optional<sheet_pr> sheet_properties_;

    optional<ext_list> extension_list_;
};

} // namespace detail
} // namespace xlnt
