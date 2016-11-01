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
struct worksheet_impl;
class xlsx_consumer;
class xlsx_producer;
}

/// <summary>
/// A worksheet is a 2D array of cells starting with cell A1 in the top-left corner
/// and extending indefinitely down and right as needed.
/// </summary>
class XLNT_API worksheet
{
public:
    using iterator = range_iterator;
    using const_iterator = const_range_iterator;
    using reverse_iterator = std::reverse_iterator<iterator>;
    using const_reverse_iterator = std::reverse_iterator<const_iterator>;

    worksheet();
    worksheet(const worksheet &rhs);

    std::string to_string() const;
    workbook &get_workbook();
	const workbook &get_workbook() const;
    void garbage_collect();

    // identification
    std::size_t get_id() const;
    void set_id(std::size_t id);

    std::string get_title() const;
    void set_title(const std::string &title);

    // freeze panes
    cell_reference get_frozen_panes() const;
    void freeze_panes(cell top_left_cell);
    void freeze_panes(const std::string &top_left_coordinate);
    void unfreeze_panes();
    bool has_frozen_panes() const;

    // container
    cell get_cell(const cell_reference &reference);
    const cell get_cell(const cell_reference &reference) const;
    bool has_cell(const cell_reference &reference) const;
    range get_range(const std::string &reference_string);
    range get_range(const range_reference &reference);
    const range get_range(const std::string &reference_string) const;
    const range get_range(const range_reference &reference) const;
    range rows() const;
    range rows(const std::string &range_string) const;
    range rows(int row_offset, int column_offset) const;
    range rows(const std::string &range_string, int row_offset, int column_offset) const;
    range columns() const;

    // properties
    column_properties &get_column_properties(column_t column);
    const column_properties &get_column_properties(column_t column) const;
    bool has_column_properties(column_t column) const;
    void add_column_properties(column_t column, const column_properties &props);

    row_properties &get_row_properties(row_t row);
    const row_properties &get_row_properties(row_t row) const;
    bool has_row_properties(row_t row) const;
    void add_row_properties(row_t row, const row_properties &props);

    // positioning
    cell_reference get_point_pos(int left, int top) const;
    cell_reference get_point_pos(const std::pair<int, int> &point) const;

    std::string unique_sheet_name(const std::string &value) const;

    // named range
    void create_named_range(const std::string &name, const std::string &reference_string);
    void create_named_range(const std::string &name, const range_reference &reference);
    bool has_named_range(const std::string &name);
    range get_named_range(const std::string &name);
    void remove_named_range(const std::string &name);

    // extents
    row_t get_lowest_row() const;
    row_t get_highest_row() const;
    row_t get_next_row() const;
    column_t get_lowest_column() const;
    column_t get_highest_column() const;
    range_reference calculate_dimension() const;
	bool has_dimension() const;

    // cell merge
    void merge_cells(const std::string &reference_string);
    void merge_cells(const range_reference &reference);
    void merge_cells(column_t start_column, row_t start_row, column_t end_column, row_t end_row);
    void unmerge_cells(const std::string &reference_string);
    void unmerge_cells(const range_reference &reference);
    void unmerge_cells(column_t start_column, row_t start_row, column_t end_column, row_t end_row);
    std::vector<range_reference> get_merged_ranges() const;

    // append
    void append();
    void append(const std::vector<std::string> &cells);
    void append(const std::vector<int> &cells);
    void append(const std::unordered_map<std::string, std::string> &cells);
    void append(const std::unordered_map<int, std::string> &cells);

    void append(const std::vector<int>::const_iterator begin, const std::vector<int>::const_iterator end);

    // operators
    bool operator==(const worksheet &other) const;
    bool operator!=(const worksheet &other) const;
    bool operator==(std::nullptr_t) const;
    bool operator!=(std::nullptr_t) const;
    void operator=(const worksheet &other);
    cell operator[](const cell_reference &reference);
    const cell operator[](const cell_reference &reference) const;
    range operator[](const range_reference &reference);
    const range operator[](const range_reference &reference) const;
    range operator[](const std::string &range_string);
    const range operator[](const std::string &range_string) const;
    range operator()(const cell_reference &top_left, const cell_reference &bottom_right);
    const range operator()(const cell_reference &top_left, const cell_reference &bottom_right) const;
    
    bool compare(const worksheet &other, bool reference) const;

    // page
	bool has_page_setup() const;
    page_setup get_page_setup() const;
	void set_page_setup(const page_setup &setup);

	bool has_page_margins() const;
    page_margins get_page_margins() const;
    void set_page_margins(const page_margins &margins);

	bool has_format_properties() const;

    // auto filter
    range_reference get_auto_filter() const;
    void auto_filter(const std::string &range_string);
    void auto_filter(const xlnt::range &range);
    void auto_filter(const range_reference &reference);
    void unset_auto_filter();
    bool has_auto_filter() const;

    // comments
    void increment_comments();
    void decrement_comments();
    std::size_t get_comment_count() const;

    void reserve(std::size_t n);

    header_footer &get_header_footer();
    const header_footer &get_header_footer() const;

    void set_parent(workbook &wb);

    std::vector<std::string> get_formula_attributes() const;

    sheet_state get_sheet_state() const;
    void set_sheet_state(sheet_state state);

    iterator begin();
    iterator end();

    const_iterator begin() const;
    const_iterator end() const;

    const_iterator cbegin() const;
    const_iterator cend() const;

    range iter_cells(bool skip_null);
    
    void set_print_title_rows(row_t first_row, row_t last_row);
    void set_print_title_rows(row_t last_row);
    void set_print_title_cols(column_t first_column, column_t last_column);
    void set_print_title_cols(column_t last_column);
    std::string get_print_titles() const;
    
    void set_print_area(const std::string &print_area);
    range_reference get_print_area() const;
    
	bool has_view() const;
    sheet_view get_view() const;

	bool x14ac_enabled() const;
	void enable_x14ac();
	void disable_x14ac();

private:
    friend class workbook;
    friend class cell;
    friend class range_iterator;
    friend class const_range_iterator;
	friend class detail::xlsx_consumer;
	friend class detail::xlsx_producer;

	void register_comments_in_manifest();

    worksheet(detail::worksheet_impl *d);
    detail::worksheet_impl *d_;
};

} // namespace xlnt
