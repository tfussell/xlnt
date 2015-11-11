// Copyright (c) 2015 Thomas Fussell
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

#include <list>
#include <memory>
#include <string>
#include <unordered_map>
#include <vector>

#include <xlnt/cell/types.hpp>
#include <xlnt/packaging/relationship.hpp>
#include <xlnt/worksheet/page_setup.hpp>

#include "xlnt_config.hpp"

namespace xlnt {

class cell;
class cell_reference;
class column_properties;
class comment;
class range;
class range_reference;
class relationship;
class row_properties;
class workbook;

struct date;

namespace detail { struct worksheet_impl; }

/// <summary>
/// Worksheet header
/// </summary>
class XLNT_CLASS header
{
  public:
    header();
    void set_text(const std::string &text)
    {
        default_ = false;
        text_ = text;
    }
    void set_font_name(const std::string &font_name)
    {
        default_ = false;
        font_name_ = font_name;
    }
    void set_font_size(std::size_t font_size)
    {
        default_ = false;
        font_size_ = font_size;
    }
    void set_font_color(const std::string &font_color)
    {
        default_ = false;
        font_color_ = font_color;
    }
    bool is_default() const
    {
        return default_;
    }

  private:
    bool default_;
    std::string text_;
    std::string font_name_;
    std::size_t font_size_;
    std::string font_color_;
};

/// <summary>
/// Worksheet footer
/// </summary>
class XLNT_CLASS footer
{
  public:
    footer();
    void set_text(const std::string &text)
    {
        default_ = false;
        text_ = text;
    }
    void set_font_name(const std::string &font_name)
    {
        default_ = false;
        font_name_ = font_name;
    }
    void set_font_size(std::size_t font_size)
    {
        default_ = false;
        font_size_ = font_size;
    }
    void set_font_color(const std::string &font_color)
    {
        default_ = false;
        font_color_ = font_color;
    }
    bool is_default() const
    {
        return default_;
    }

  private:
    bool default_;
    std::string text_;
    std::string font_name_;
    std::size_t font_size_;
    std::string font_color_;
};

/// <summary>
/// Worksheet header and footer
/// </summary>
class XLNT_CLASS header_footer
{
  public:
    header_footer();

    header &get_left_header()
    {
        return left_header_;
    }
    header &get_center_header()
    {
        return center_header_;
    }
    header &get_right_header()
    {
        return right_header_;
    }
    footer &get_left_footer()
    {
        return left_footer_;
    }
    footer &get_center_footer()
    {
        return center_footer_;
    }
    footer &get_right_footer()
    {
        return right_footer_;
    }

    bool is_default_header() const
    {
        return left_header_.is_default() && center_header_.is_default() && right_header_.is_default();
    }
    bool is_default_footer() const
    {
        return left_footer_.is_default() && center_footer_.is_default() && right_footer_.is_default();
    }
    bool is_default() const
    {
        return is_default_header() && is_default_footer();
    }

  private:
    header left_header_, right_header_, center_header_;
    footer left_footer_, right_footer_, center_footer_;
};

/// <summary>
/// Worksheet margins
/// </summary>
struct XLNT_CLASS margins
{
  public:
    margins() : default_(true), top_(0), left_(0), bottom_(0), right_(0), header_(0), footer_(0)
    {
    }

    bool is_default() const
    {
        return default_;
    }
    double get_top() const
    {
        return top_;
    }
    void set_top(double top)
    {
        default_ = false;
        top_ = top;
    }
    double get_left() const
    {
        return left_;
    }
    void set_left(double left)
    {
        default_ = false;
        left_ = left;
    }
    double get_bottom() const
    {
        return bottom_;
    }
    void set_bottom(double bottom)
    {
        default_ = false;
        bottom_ = bottom;
    }
    double get_right() const
    {
        return right_;
    }
    void set_right(double right)
    {
        default_ = false;
        right_ = right;
    }
    double get_header() const
    {
        return header_;
    }
    void set_header(double header)
    {
        default_ = false;
        header_ = header;
    }
    double get_footer() const
    {
        return footer_;
    }
    void set_footer(double footer)
    {
        default_ = false;
        footer_ = footer;
    }

  private:
    bool default_;
    double top_;
    double left_;
    double bottom_;
    double right_;
    double header_;
    double footer_;
};

/// <summary>
/// A worksheet is a 2D array of cells.
/// </summary>
class XLNT_CLASS worksheet
{
  public:
    worksheet();
    worksheet(const worksheet &rhs);
    worksheet(workbook &parent_workbook, const std::string &title = std::string());

    std::string to_string() const;
    workbook &get_parent() const;
    void garbage_collect();

    // title
    std::string get_title() const;
    void set_title(const std::string &title);
    std::string make_unique_sheet_name(const std::string &value);

    // freeze panes
    cell_reference get_frozen_panes() const;
    void freeze_panes(cell top_left_cell);
    void freeze_panes(const std::string &top_left_coordinate);
    void unfreeze_panes();
    bool has_frozen_panes() const;

    // container
    cell get_cell(const cell_reference &reference);
    const cell get_cell(const cell_reference &reference) const;
    range get_range(const std::string &reference_string);
    range get_range(const range_reference &reference);
    const range get_range(const std::string &reference_string) const;
    const range get_range(const range_reference &reference) const;
    range get_squared_range(column_t min_col, row_t min_row, column_t max_col, row_t max_row);
    const range get_squared_range(column_t min_col, row_t min_row, column_t max_col, row_t max_row) const;
    range rows() const;
    range rows(const std::string &range_string) const;
    range rows(int row_offset, int column_offset) const;
    range rows(const std::string &range_string, int row_offset, int column_offset) const;
    range columns() const;
    std::list<cell> get_cell_collection();

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

    // relationships
    relationship create_relationship(relationship::type type, const std::string &target_uri);
    const std::vector<relationship> &get_relationships() const;

    // charts
    // void add_chart(chart chart);

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
    void append(const std::vector<date> &cells);
    void append(const std::vector<cell> &cells);
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

    // page
    page_setup &get_page_setup();
    const page_setup &get_page_setup() const;
    margins &get_page_margins();
    const margins &get_page_margins() const;

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

    void set_sheet_state(page_setup::sheet_state state);

  private:
    friend class workbook;
    friend class cell;
    worksheet(detail::worksheet_impl *d);
    detail::worksheet_impl *d_;
};

} // namespace xlnt
