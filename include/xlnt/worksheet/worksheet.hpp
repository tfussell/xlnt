// Copyright (c) 2014-2017 Thomas Fussell
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

#pragma once

#include <iterator>
#include <memory>
#include <string>
#include <unordered_map>
#include <vector>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/cell/index_types.hpp>
#include <xlnt/packaging/relationship.hpp>
#include <xlnt/worksheet/page_margins.hpp>
#include <xlnt/worksheet/page_setup.hpp>
#include <xlnt/worksheet/sheet_view.hpp>

namespace xlnt {

class cell;
class cell_reference;
class cell_vector;
class column_properties;
class comment;
class const_range_iterator;
class footer;
class header;
class range;
class range_iterator;
class range_reference;
class relationship;
class row_properties;
class workbook;

struct date;

namespace detail {

class xlsx_consumer;
class xlsx_producer;

struct worksheet_impl;

} // namespace detail

/// <summary>
/// A worksheet is a 2D array of cells starting with cell A1 in the top-left corner
/// and extending indefinitely down and right as needed.
/// </summary>
class XLNT_API worksheet
{
public:
    /// <summary>
    /// Iterate over a non-const worksheet with an iterator of this type.
    /// </summary>
    using iterator = range_iterator;

    /// <summary>
    /// Iterate over a non-const worksheet with an iterator of this type.
    /// </summary>
    using const_iterator = const_range_iterator;

    /// <summary>
    /// Iterate in reverse over a non-const worksheet with an iterator of this type.
    /// </summary>
    using reverse_iterator = std::reverse_iterator<iterator>;

    /// <summary>
    /// Iterate in reverse order over a const worksheet with an iterator of this type.
    /// </summary>
    using const_reverse_iterator = std::reverse_iterator<const_iterator>;

    /// <summary>
    /// Construct a null worksheet. No methods should be called on such a worksheet.
    /// </summary>
    worksheet();

    /// <summary>
    /// Copy constructor. This worksheet will point to the same memory as rhs's worksheet.
    /// </summary>
    worksheet(const worksheet &rhs);

    /// <summary>
    /// Returns a reference to the workbook this worksheet is owned by.
    /// </summary>
    class workbook &workbook();

    /// <summary>
    /// Returns a reference to the workbook this worksheet is owned by.
    /// </summary>
    const class workbook &workbook() const;

    /// <summary>
    /// Deletes data held in the worksheet that does not affect the internal data or display.
    /// For example, unreference styles and empty cells will be removed.
    /// </summary>
    void garbage_collect();

    // identification

    /// <summary>
    /// Returns the unique numeric identifier of this worksheet. This will sometimes but not necessarily
    /// be the index of the worksheet in the workbook.
    /// </summary>
    std::size_t id() const;

    /// <summary>
    /// Set the unique numeric identifier. The id defaults to the lowest unused id in the workbook
    /// so this should not be called without a good reason.
    /// </summary>
    void id(std::size_t id);

    /// <summary>
    /// Returns the title of this sheet.
    /// </summary>
    std::string title() const;

    /// <summary>
    /// Sets the title of this sheet.
    /// </summary>
    void title(const std::string &title);

    // freeze panes

    /// <summary>
    /// Returns the top left corner of the region above and to the left of which panes are frozen.
    /// </summary>
    cell_reference frozen_panes() const;

    /// <summary>
    /// Freeze panes above and to the left of top_left_cell.
    /// </summary>
    void freeze_panes(cell top_left_cell);

    /// <summary>
    /// Freeze panes above and to the left of top_left_coordinate.
    /// </summary>
    void freeze_panes(const cell_reference &top_left_coordinate);

    /// <summary>
    /// Remove frozen panes. The data in those panes will be unaffected--this affects only the view.
    /// </summary>
    void unfreeze_panes();

    /// <summary>
    /// Returns true if this sheet has a frozen row, frozen column, or both.
    /// </summary>
    bool has_frozen_panes() const;

    // container

    /// <summary>
    ///
    /// </summary>
    class cell cell(column_t column, row_t row);

    /// <summary>
    ///
    /// </summary>
    const class cell cell(column_t column, row_t row) const;

    /// <summary>
    ///
    /// </summary>
    class cell cell(const cell_reference &reference);

    /// <summary>
    ///
    /// </summary>
    const class cell cell(const cell_reference &reference) const;

    /// <summary>
    ///
    /// </summary>
    bool has_cell(const cell_reference &reference) const;

    /// <summary>
    ///
    /// </summary>
    class range range(const std::string &reference_string);

    /// <summary>
    ///
    /// </summary>
    class range range(const range_reference &reference);

    /// <summary>
    ///
    /// </summary>
    const class range range(const std::string &reference_string) const;

    /// <summary>
    ///
    /// </summary>
    const class range range(const range_reference &reference) const;

    /// <summary>
    ///
    /// </summary>
    class range rows() const;

    /// <summary>
    ///
    /// </summary>
    class range rows(const std::string &range_string) const;

    /// <summary>
    ///
    /// </summary>
    class range rows(int row_offset, int column_offset) const;

    /// <summary>
    ///
    /// </summary>
    class range rows(const std::string &range_string, int row_offset, int column_offset) const;

    /// <summary>
    ///
    /// </summary>
    class range columns() const;

    // properties

    /// <summary>
    ///
    /// </summary>
    xlnt::column_properties &column_properties(column_t column);

    /// <summary>
    ///
    /// </summary>
    const xlnt::column_properties &column_properties(column_t column) const;

    /// <summary>
    ///
    /// </summary>
    bool has_column_properties(column_t column) const;

    /// <summary>
    ///
    /// </summary>
    void add_column_properties(column_t column, const class column_properties &props);

    /// <summary>
    /// Calculate the width of the given column. This will be the default column width if
    /// a custom width is not set on this column's column_properties.
    /// </summary>
    double column_width(column_t column) const;

    /// <summary>
    ///
    /// </summary>
    xlnt::row_properties &row_properties(row_t row);

    /// <summary>
    ///
    /// </summary>
    const xlnt::row_properties &row_properties(row_t row) const;

    /// <summary>
    ///
    /// </summary>
    bool has_row_properties(row_t row) const;

    /// <summary>
    ///
    /// </summary>
    void add_row_properties(row_t row, const class row_properties &props);

    /// <summary>
    /// Calculate the height of the given row. This will be the default row height if
    /// a custom height is not set on this row's row_properties.
    /// </summary>
    double row_height(row_t row) const;

    // positioning

    /// <summary>
    ///
    /// </summary>
    cell_reference point_pos(int left, int top) const;

    /// <summary>
    ///
    /// </summary>
    cell_reference point_pos(const std::pair<int, int> &point) const;

    /// <summary>
    ///
    /// </summary>
    std::string unique_sheet_name(const std::string &value) const;

    // named range

    /// <summary>
    ///
    /// </summary>
    void create_named_range(const std::string &name, const std::string &reference_string);

    /// <summary>
    ///
    /// </summary>
    void create_named_range(const std::string &name, const range_reference &reference);

    /// <summary>
    ///
    /// </summary>
    bool has_named_range(const std::string &name);

    /// <summary>
    ///
    /// </summary>
    class range named_range(const std::string &name);

    /// <summary>
    ///
    /// </summary>
    void remove_named_range(const std::string &name);

    // extents

    /// <summary>
    ///
    /// </summary>
    row_t lowest_row() const;

    /// <summary>
    ///
    /// </summary>
    row_t highest_row() const;

    /// <summary>
    ///
    /// </summary>
    row_t next_row() const;

    /// <summary>
    ///
    /// </summary>
    column_t lowest_column() const;

    /// <summary>
    ///
    /// </summary>
    column_t highest_column() const;

    /// <summary>
    ///
    /// </summary>
    range_reference calculate_dimension() const;

    // cell merge

    /// <summary>
    ///
    /// </summary>
    void merge_cells(const std::string &reference_string);

    /// <summary>
    ///
    /// </summary>
    void merge_cells(const range_reference &reference);

    /// <summary>
    ///
    /// </summary>
    void merge_cells(column_t start_column, row_t start_row, column_t end_column, row_t end_row);

    /// <summary>
    ///
    /// </summary>
    void unmerge_cells(const std::string &reference_string);

    /// <summary>
    ///
    /// </summary>
    void unmerge_cells(const range_reference &reference);

    /// <summary>
    ///
    /// </summary>
    void unmerge_cells(column_t start_column, row_t start_row, column_t end_column, row_t end_row);

    /// <summary>
    ///
    /// </summary>
    std::vector<range_reference> merged_ranges() const;

    // append

    /// <summary>
    ///
    /// </summary>
    void append();

    /// <summary>
    ///
    /// </summary>
    void append(const std::vector<std::string> &cells);

    /// <summary>
    ///
    /// </summary>
    void append(const std::vector<int> &cells);

    /// <summary>
    ///
    /// </summary>
    void append(const std::unordered_map<std::string, std::string> &cells);

    /// <summary>
    ///
    /// </summary>
    void append(const std::unordered_map<int, std::string> &cells);

    /// <summary>
    ///
    /// </summary>
    void append(const std::vector<int>::const_iterator begin, const std::vector<int>::const_iterator end);

    // operators

    /// <summary>
    ///
    /// </summary>
    bool operator==(const worksheet &other) const;

    /// <summary>
    ///
    /// </summary>
    bool operator!=(const worksheet &other) const;

    /// <summary>
    ///
    /// </summary>
    bool operator==(std::nullptr_t) const;

    /// <summary>
    ///
    /// </summary>
    bool operator!=(std::nullptr_t) const;

    /// <summary>
    ///
    /// </summary>
    void operator=(const worksheet &other);

    /// <summary>
    ///
    /// </summary>
    class cell operator[](const cell_reference &reference);

    /// <summary>
    ///
    /// </summary>
    const class cell operator[](const cell_reference &reference) const;

    /// <summary>
    ///
    /// </summary>
    class range operator[](const range_reference &reference);

    /// <summary>
    ///
    /// </summary>
    const class range operator[](const range_reference &reference) const;

    /// <summary>
    ///
    /// </summary>
    class range operator[](const std::string &range_string);

    /// <summary>
    ///
    /// </summary>
    const class range operator[](const std::string &range_string) const;

    /// <summary>
    ///
    /// </summary>
    class range operator()(const cell_reference &top_left, const cell_reference &bottom_right);

    /// <summary>
    ///
    /// </summary>
    const class range operator()(const cell_reference &top_left, const cell_reference &bottom_right) const;

    /// <summary>
    ///
    /// </summary>
    bool compare(const worksheet &other, bool reference) const;

    // page

    /// <summary>
    ///
    /// </summary>
    bool has_page_setup() const;

    /// <summary>
    ///
    /// </summary>
    xlnt::page_setup page_setup() const;

    /// <summary>
    ///
    /// </summary>
    void page_setup(const struct page_setup &setup);

    /// <summary>
    ///
    /// </summary>
    bool has_page_margins() const;

    /// <summary>
    ///
    /// </summary>
    xlnt::page_margins page_margins() const;

    /// <summary>
    ///
    /// </summary>
    void page_margins(const class page_margins &margins);

    // auto filter

    /// <summary>
    ///
    /// </summary>
    range_reference auto_filter() const;

    /// <summary>
    ///
    /// </summary>
    void auto_filter(const std::string &range_string);

    /// <summary>
    ///
    /// </summary>
    void auto_filter(const xlnt::range &range);

    /// <summary>
    ///
    /// </summary>
    void auto_filter(const range_reference &reference);

    /// <summary>
    ///
    /// </summary>
    void clear_auto_filter();

    /// <summary>
    ///
    /// </summary>
    bool has_auto_filter() const;

    /// <summary>
    ///
    /// </summary>
    void reserve(std::size_t n);

    /// <summary>
    ///
    /// </summary>
    bool has_header_footer() const;

    /// <summary>
    ///
    /// </summary>
    class header_footer header_footer() const;

    /// <summary>
    ///
    /// </summary>
    void header_footer(const class header_footer &new_header_footer);

    /// <summary>
    ///
    /// </summary>
    void parent(class workbook &wb);

    /// <summary>
    ///
    /// </summary>
    std::vector<std::string> formula_attributes() const;

    /// <summary>
    ///
    /// </summary>
    xlnt::sheet_state sheet_state() const;

    /// <summary>
    ///
    /// </summary>
    void sheet_state(xlnt::sheet_state state);

    /// <summary>
    ///
    /// </summary>
    iterator begin();

    /// <summary>
    ///
    /// </summary>
    iterator end();

    /// <summary>
    ///
    /// </summary>
    const_iterator begin() const;

    /// <summary>
    ///
    /// </summary>
    const_iterator end() const;

    /// <summary>
    ///
    /// </summary>
    const_iterator cbegin() const;

    /// <summary>
    ///
    /// </summary>
    const_iterator cend() const;

    /// <summary>
    ///
    /// </summary>
    class range iter_cells(bool skip_null);

    /// <summary>
    ///
    /// </summary>
    void print_title_rows(row_t first_row, row_t last_row);

    /// <summary>
    ///
    /// </summary>
    void print_title_rows(row_t last_row);

    /// <summary>
    ///
    /// </summary>
    void print_title_cols(column_t first_column, column_t last_column);

    /// <summary>
    ///
    /// </summary>
    void print_title_cols(column_t last_column);

    /// <summary>
    ///
    /// </summary>
    std::string print_titles() const;

    /// <summary>
    ///
    /// </summary>
    void print_area(const std::string &print_area);

    /// <summary>
    ///
    /// </summary>
    range_reference print_area() const;

    /// <summary>
    ///
    /// </summary>
    bool has_view() const;

    /// <summary>
    ///
    /// </summary>
    sheet_view view(std::size_t index = 0) const;

    /// <summary>
    ///
    /// </summary>
    void add_view(const sheet_view &new_view);

    // page breaks

    /// <summary>
    /// Remove all manual column and row page breaks (represented as dashed
    /// blue lines in the page view in Excel).
    /// </summary>
    void clear_page_breaks();

    /// <summary>
    /// Returns vector where each element represents a row which will break a page below it.
    /// </summary>
    const std::vector<row_t> &page_break_rows() const;

    /// <summary>
    /// Add a page break at the given row.
    /// </summary>
    void page_break_at_row(row_t row);

    /// <summary>
    /// Returns vector where each element represents a column which will break a page to the right.
    /// </summary>
    const std::vector<column_t> &page_break_columns() const;

    /// <summary>
    /// Add a page break at the given column.
    /// </summary>
    void page_break_at_column(column_t column);

private:
    friend class cell;
    friend class const_range_iterator;
    friend class range_iterator;
    friend class workbook;
    friend class detail::xlsx_consumer;
    friend class detail::xlsx_producer;

    /// <summary>
    ///
    /// </summary>
    void register_comments_in_manifest();

    /// <summary>
    /// Removes calcChain part from manifest if no formulae remain in workbook.
    /// </summary>
    void garbage_collect_formulae();

    /// <summary>
    ///
    /// </summary>
    worksheet(detail::worksheet_impl *d);

    /// <summary>
    ///
    /// </summary>
    detail::worksheet_impl *d_;
};

} // namespace xlnt
