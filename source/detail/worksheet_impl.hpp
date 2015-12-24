// Copyright (c) 2014-2016 Thomas Fussell
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
#include <string>
#include <unordered_map>
#include <vector>

#include <xlnt/worksheet/column_properties.hpp>
#include <xlnt/worksheet/row_properties.hpp>

#include "cell_impl.hpp"

namespace xlnt {

class workbook;

namespace detail {

struct worksheet_impl
{
    worksheet_impl(workbook *parent_workbook, const std::string &title)
        : parent_(parent_workbook), title_(title), freeze_panes_("A1"), comment_count_(0)
    {
        page_margins_.set_left(0.75);
        page_margins_.set_right(0.75);
        page_margins_.set_top(1);
        page_margins_.set_bottom(1);
        page_margins_.set_header(0.5);
        page_margins_.set_footer(0.5);
    }

    worksheet_impl(const worksheet_impl &other)
    {
        *this = other;
    }

    void operator=(const worksheet_impl &other)
    {
        parent_ = other.parent_;
        column_properties_ = other.column_properties_;
        row_properties_ = other.row_properties_;
        title_ = other.title_;
        freeze_panes_ = other.freeze_panes_;
        cell_map_ = other.cell_map_;
        for (auto &row : cell_map_)
        {
            for (auto &cell : row.second)
            {
                cell.second.parent_ = this;
            }
        }
        relationships_ = other.relationships_;
        page_setup_ = other.page_setup_;
        auto_filter_ = other.auto_filter_;
        page_margins_ = other.page_margins_;
        merged_cells_ = other.merged_cells_;
        named_ranges_ = other.named_ranges_;
        comment_count_ = other.comment_count_;
        header_footer_ = other.header_footer_;
    }

    workbook *parent_;
    std::unordered_map<column_t, column_properties> column_properties_;
    std::unordered_map<row_t, row_properties> row_properties_;
    std::string title_;
    cell_reference freeze_panes_;
    std::unordered_map<row_t, std::unordered_map<column_t, cell_impl>> cell_map_;
    std::vector<relationship> relationships_;
    page_setup page_setup_;
    range_reference auto_filter_;
    page_margins page_margins_;
    std::vector<range_reference> merged_cells_;
    std::unordered_map<std::string, named_range> named_ranges_;
    std::size_t comment_count_;
    header_footer header_footer_;
};

} // namespace detail
} // namespace xlnt
