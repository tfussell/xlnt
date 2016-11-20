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
#pragma once

#include <iterator>
#include <memory>
#include <string>
#include <unordered_map>
#include <vector>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/cell/index_types.hpp>
#include <xlnt/packaging/relationship.hpp>
#include <xlnt/worksheet/header_footer.hpp>
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

}

/// <summary>
/// A worksheet is a 2D array of cells starting with cell A1 in the top-left corner
/// and extending indefinitely down and right as needed.
/// </summary>
class XLNT_API worksheet
{
public:
    /// <summary>
    ///
    /// </summary>
    using iterator = range_iterator;

    /// <summary>
    ///
    /// </summary>
    using const_iterator = const_range_iterator;

    /// <summary>
    ///
    /// </summary>
    using reverse_iterator = std::reverse_iterator<iterator>;

    /// <summary>
    ///
    /// </summary>
    using const_reverse_iterator = std::reverse_iterator<const_iterator>;

    /// <summary>
    ///
    /// </summary>
    worksheet();

    /// <summary>
    ///
    /// </summary>
    worksheet(const worksheet &rhs);

    /// <summary>
    ///
    /// </summary>
    std::string to_string() const;

    /// <summary>
    ///
    /// </summary>
    workbook &get_workbook();

    /// <summary>
    ///
    /// </summary>
	const workbook &get_workbook() const;

    /// <summary>
    ///
    /// </summary>
    void garbage_collect();

    // identification

    /// <summary>
    ///
    /// </summary>
    std::size_t get_id() const;

    /// <summary>
    ///
    /// </summary>
    void set_id(std::size_t id);

    /// <summary>
    ///
    /// </summary>
    std::string get_title() const;

    /// <summary>
    ///
    /// </summary>
    void set_title(const std::string &title);

    // freeze panes

    /// <summary>
    ///
    /// </summary>
    cell_reference get_frozen_panes() const;

    /// <summary>
    ///
    /// </summary>
    void freeze_panes(cell top_left_cell);

    /// <summary>
    ///
    /// </summary>
    void freeze_panes(const std::string &top_left_coordinate);

    /// <summary>
    ///
    /// </summary>
    void unfreeze_panes();

    /// <summary>
    ///
    /// </summary>
    bool has_frozen_panes() const;

    // container

    /// <summary>
    ///
    /// </summary>
    cell get_cell(column_t column, row_t row);

    /// <summary>
    ///
    /// </summary>
    const cell get_cell(column_t column, row_t row) const;

    /// <summary>
    ///
    /// </summary>
    cell get_cell(const cell_reference &reference);

    /// <summary>
    ///
    /// </summary>
    const cell get_cell(const cell_reference &reference) const;

    /// <summary>
    ///
    /// </summary>
    bool has_cell(const cell_reference &reference) const;

    /// <summary>
    ///
    /// </summary>
    range get_range(const std::string &reference_string);

    /// <summary>
    ///
    /// </summary>
    range get_range(const range_reference &reference);

    /// <summary>
    ///
    /// </summary>
    const range get_range(const std::string &reference_string) const;

    /// <summary>
    ///
    /// </summary>
    const range get_range(const range_reference &reference) const;

    /// <summary>
    ///
    /// </summary>
    range rows() const;

    /// <summary>
    ///
    /// </summary>
    range rows(const std::string &range_string) const;

    /// <summary>
    ///
    /// </summary>
    range rows(int row_offset, int column_offset) const;

    /// <summary>
    ///
    /// </summary>
    range rows(const std::string &range_string, int row_offset, int column_offset) const;

    /// <summary>
    ///
    /// </summary>
    range columns() const;

    // properties

    /// <summary>
    ///
    /// </summary>
    column_properties &get_column_properties(column_t column);

    /// <summary>
    ///
    /// </summary>
    const column_properties &get_column_properties(column_t column) const;

    /// <summary>
    ///
    /// </summary>
    bool has_column_properties(column_t column) const;

    /// <summary>
    ///
    /// </summary>
    void add_column_properties(column_t column, const column_properties &props);

    /// <summary>
    ///
    /// </summary>
    row_properties &get_row_properties(row_t row);

    /// <summary>
    ///
    /// </summary>
    const row_properties &get_row_properties(row_t row) const;

    /// <summary>
    ///
    /// </summary>
    bool has_row_properties(row_t row) const;

    /// <summary>
    ///
    /// </summary>
    void add_row_properties(row_t row, const row_properties &props);

    // positioning

    /// <summary>
    ///
    /// </summary>
    cell_reference get_point_pos(int left, int top) const;

    /// <summary>
    ///
    /// </summary>
    cell_reference get_point_pos(const std::pair<int, int> &point) const;

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
    range get_named_range(const std::string &name);

    /// <summary>
    ///
    /// </summary>
    void remove_named_range(const std::string &name);

    // extents

    /// <summary>
    ///
    /// </summary>
    row_t get_lowest_row() const;

    /// <summary>
    ///
    /// </summary>
    row_t get_highest_row() const;

    /// <summary>
    ///
    /// </summary>
    row_t get_next_row() const;

    /// <summary>
    ///
    /// </summary>
    column_t get_lowest_column() const;

    /// <summary>
    ///
    /// </summary>
    column_t get_highest_column() const;

    /// <summary>
    ///
    /// </summary>
    range_reference calculate_dimension() const;

    /// <summary>
    ///
    /// </summary>
	bool has_dimension() const;

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
    std::vector<range_reference> get_merged_ranges() const;

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
    cell operator[](const cell_reference &reference);

    /// <summary>
    ///
    /// </summary>
    const cell operator[](const cell_reference &reference) const;

    /// <summary>
    ///
    /// </summary>
    range operator[](const range_reference &reference);

    /// <summary>
    ///
    /// </summary>
    const range operator[](const range_reference &reference) const;

    /// <summary>
    ///
    /// </summary>
    range operator[](const std::string &range_string);

    /// <summary>
    ///
    /// </summary>
    const range operator[](const std::string &range_string) const;

    /// <summary>
    ///
    /// </summary>
    range operator()(const cell_reference &top_left, const cell_reference &bottom_right);

    /// <summary>
    ///
    /// </summary>
    const range operator()(const cell_reference &top_left, const cell_reference &bottom_right) const;

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
    page_setup get_page_setup() const;

    /// <summary>
    ///
    /// </summary>
	void set_page_setup(const page_setup &setup);

    /// <summary>
    ///
    /// </summary>
	bool has_page_margins() const;

    /// <summary>
    ///
    /// </summary>
    page_margins get_page_margins() const;

    /// <summary>
    ///
    /// </summary>
    void set_page_margins(const page_margins &margins);

    /// <summary>
    ///
    /// </summary>
	bool has_format_properties() const;

    // auto filter

    /// <summary>
    ///
    /// </summary>
    range_reference get_auto_filter() const;

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
    void unset_auto_filter();

    /// <summary>
    ///
    /// </summary>
    bool has_auto_filter() const;

    // comments

    /// <summary>
    ///
    /// </summary>
    void increment_comments();

    /// <summary>
    ///
    /// </summary>
    void decrement_comments();

    /// <summary>
    ///
    /// </summary>
    std::size_t get_comment_count() const;

    /// <summary>
    ///
    /// </summary>
    void reserve(std::size_t n);

    /// <summary>
    ///
    /// </summary>
    header_footer &get_header_footer();

    /// <summary>
    ///
    /// </summary>
    const header_footer &get_header_footer() const;

    /// <summary>
    ///
    /// </summary>
    void set_parent(workbook &wb);

    /// <summary>
    ///
    /// </summary>
    std::vector<std::string> get_formula_attributes() const;

    /// <summary>
    ///
    /// </summary>
    sheet_state get_sheet_state() const;

    /// <summary>
    ///
    /// </summary>
    void set_sheet_state(sheet_state state);

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
    range iter_cells(bool skip_null);

    /// <summary>
    ///
    /// </summary>
    void set_print_title_rows(row_t first_row, row_t last_row);

    /// <summary>
    ///
    /// </summary>
    void set_print_title_rows(row_t last_row);

    /// <summary>
    ///
    /// </summary>
    void set_print_title_cols(column_t first_column, column_t last_column);

    /// <summary>
    ///
    /// </summary>
    void set_print_title_cols(column_t last_column);

    /// <summary>
    ///
    /// </summary>
    std::string get_print_titles() const;

    /// <summary>
    ///
    /// </summary>
    void set_print_area(const std::string &print_area);

    /// <summary>
    ///
    /// </summary>
    range_reference get_print_area() const;

    /// <summary>
    ///
    /// </summary>
	bool has_view() const;

    /// <summary>
    ///
    /// </summary>
    sheet_view get_view() const;

    /// <summary>
    ///
    /// </summary>
	bool x14ac_enabled() const;

    /// <summary>
    ///
    /// </summary>
	void enable_x14ac();

    /// <summary>
    ///
    /// </summary>
	void disable_x14ac();

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
    ///
    /// </summary>
    worksheet(detail::worksheet_impl *d);

    /// <summary>
    ///
    /// </summary>
    detail::worksheet_impl *d_;
};

} // namespace xlnt
