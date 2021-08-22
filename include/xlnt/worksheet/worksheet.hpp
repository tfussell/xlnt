// Copyright (c) 2014-2021 Thomas Fussell
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
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
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
class condition;
class conditional_format;
class const_range_iterator;
class footer;
class header;
class range;
class range_iterator;
class range_reference;
class relationship;
class row_properties;
class sheet_format_properties;
class workbook;
class phonetic_pr;

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
    /// Returns true if this sheet has an initialized cell at the given reference.
    /// </summary>
    bool has_cell(const cell_reference &reference) const;

    /// <summary>
    /// Returns the cell at the given reference. If the cell doesn't exist, it
    /// will be initialized to null before being returned.
    /// </summary>
    class cell cell(const cell_reference &reference);

    /// <summary>
    /// Returns the cell at the given reference. If the cell doesn't exist, an
    /// invalid_parameter exception will be thrown.
    /// </summary>
    const class cell cell(const cell_reference &reference) const;

    /// <summary>
    /// Returns the cell at the given column and row. If the cell doesn't exist, it
    /// will be initialized to null before being returned.
    /// </summary>
    class cell cell(column_t column, row_t row);

    /// <summary>
    /// Returns the cell at the given column and row. If the cell doesn't exist, an
    /// invalid_parameter exception will be thrown.
    /// </summary>
    const class cell cell(column_t column, row_t row) const;

    /// <summary>
    /// Returns the range defined by reference string. If reference string is the name of
    /// a previously-defined named range in the sheet, it will be returned.
    /// </summary>
    class range range(const std::string &reference_string);

    /// <summary>
    /// Returns the range defined by reference string. If reference string is the name of
    /// a previously-defined named range in the sheet, it will be returned.
    /// </summary>
    const class range range(const std::string &reference_string) const;

    /// <summary>
    /// Returns the range defined by reference.
    /// </summary>
    class range range(const range_reference &reference);

    /// <summary>
    /// Returns the range defined by reference.
    /// </summary>
    const class range range(const range_reference &reference) const;

    /// <summary>
    /// Returns a range encompassing all cells in this sheet which will
    /// be iterated upon in row-major order. If skip_null is true (default),
    /// empty rows and cells will be skipped during iteration of the range.
    /// </summary>
    class range rows(bool skip_null = true);

    /// <summary>
    /// Returns a range encompassing all cells in this sheet which will
    /// be iterated upon in row-major order. If skip_null is true (default),
    /// empty rows and cells will be skipped during iteration of the range.
    /// </summary>
    const class range rows(bool skip_null = true) const;

    /// <summary>
    /// Returns a range ecompassing all cells in this sheet which will
    /// be iterated upon in column-major order. If skip_null is true (default),
    /// empty columns and cells will be skipped during iteration of the range.
    /// </summary>
    class range columns(bool skip_null = true);

    /// <summary>
    /// Returns a range ecompassing all cells in this sheet which will
    /// be iterated upon in column-major order. If skip_null is true (default),
    /// empty columns and cells will be skipped during iteration of the range.
    /// </summary>
    const class range columns(bool skip_null = true) const;

    //TODO: finish implementing cell_iterator wrapping before uncommenting
    //class cell_vector cells(bool skip_null = true);

    //TODO: finish implementing cell_iterator wrapping before uncommenting
    //const class cell_vector cells(bool skip_null = true) const;

    /// <summary>
    /// Clears memory used by the given cell.
    /// </summary>
    void clear_cell(const cell_reference &ref);

    /// <summary>
    /// Clears memory used by all cells in the given row.
    /// </summary>
    void clear_row(row_t row);

    /// <summary>
    /// Insert empty rows before the given row index
    /// </summary>
    void insert_rows(row_t row, std::uint32_t amount);

    /// <summary>
    /// Insert empty columns before the given column index
    /// </summary>
    void insert_columns(column_t column, std::uint32_t amount);

    /// <summary>
    /// Delete rows before the given row index
    /// </summary>
    void delete_rows(row_t row, std::uint32_t amount);

    /// <summary>
    /// Delete columns before the given column index
    /// </summary>
    void delete_columns(column_t column, std::uint32_t amount);

    // properties

    /// <summary>
    /// Returns the column properties for the given column.
    /// </summary>
    xlnt::column_properties &column_properties(column_t column);

    /// <summary>
    /// Returns the column properties for the given column.
    /// </summary>
    const xlnt::column_properties &column_properties(column_t column) const;

    /// <summary>
    /// Returns true if column properties have been set for the given column.
    /// </summary>
    bool has_column_properties(column_t column) const;

    /// <summary>
    /// Sets column properties for the given column to props.
    /// </summary>
    void add_column_properties(column_t column, const class column_properties &props);

    /// <summary>
    /// Calculates the width of the given column. This will be the default column width if
    /// a custom width is not set on this column's column_properties.
    /// </summary>
    double column_width(column_t column) const;

    /// <summary>
    /// Returns the row properties for the given row.
    /// </summary>
    xlnt::row_properties &row_properties(row_t row);

    /// <summary>
    /// Returns the row properties for the given row.
    /// </summary>
    const xlnt::row_properties &row_properties(row_t row) const;

    /// <summary>
    /// Returns true if row properties have been set for the given row.
    /// </summary>
    bool has_row_properties(row_t row) const;

    /// <summary>
    /// Sets row properties for the given row to props.
    /// </summary>
    void add_row_properties(row_t row, const class row_properties &props);

    /// <summary>
    /// Calculate the height of the given row. This will be the default row height if
    /// a custom height is not set on this row's row_properties.
    /// </summary>
    double row_height(row_t row) const;

    // positioning

    /// <summary>
    /// Returns a reference to the cell at the given point coordinates.
    /// </summary>
    cell_reference point_pos(int left, int top) const;

    // named range

    /// <summary>
    /// Creates a new named range with the given name encompassing the string representing a range.
    /// </summary>
    void create_named_range(const std::string &name, const std::string &reference_string);

    /// <summary>
    /// Creates a new named range with the given name encompassing the given range reference.
    /// </summary>
    void create_named_range(const std::string &name, const range_reference &reference);

    /// <summary>
    /// Returns true if this worksheet contains a named range with the given name.
    /// </summary>
    bool has_named_range(const std::string &name) const;

    /// <summary>
    /// Returns the named range with the given name. Throws key_not_found
    /// exception if the named range doesn't exist.
    /// </summary>
    class range named_range(const std::string &name);

    /// <summary>
    /// Returns the named range with the given name. Throws key_not_found
    /// exception if the named range doesn't exist.
    /// </summary>
    const class range named_range(const std::string &name) const;

    /// <summary>
    /// Removes a named range with the given name.
    /// </summary>
    void remove_named_range(const std::string &name);

    // extents

    /// <summary>
    /// Returns the row of the first non-empty cell in the worksheet.
    /// </summary>
    row_t lowest_row() const;

    /// <summary>
    /// Returns the row of the first non-empty cell or lowest row with properties in the worksheet.
    /// </summary>
    row_t lowest_row_or_props() const;

    /// <summary>
    /// Returns the row of the last non-empty cell in the worksheet.
    /// </summary>
    row_t highest_row() const;

    /// <summary>
    /// Returns the row of the last non-empty cell or highest row with properties in the worksheet.
    /// </summary>
    row_t highest_row_or_props() const;

    /// <summary>
    /// Returns the row directly below the last non-empty cell in the worksheet.
    /// </summary>
    row_t next_row() const;

    /// <summary>
    /// Returns the column of the first non-empty cell in the worksheet.
    /// </summary>
    column_t lowest_column() const;

    /// <summary>
    /// Returns the column of the first non-empty cell or lowest column with properties in the worksheet.
    /// </summary>
    column_t lowest_column_or_props() const;

    /// <summary>
    /// Returns the column of the last non-empty cell in the worksheet.
    /// </summary>
    column_t highest_column() const;

    /// <summary>
    /// Returns the column of the last non-empty cell or highest column with properties in the worksheet.
    /// </summary>
    column_t highest_column_or_props() const;

    /// <summary>
    /// Returns a range_reference pointing to the full range of cells in the worksheet.
    /// If skip_null is true (default), empty cells (For example if the first row or column is empty)
    /// will not be included in upper left bound of range.
    /// </summary>
    range_reference calculate_dimension(bool skip_null=true) const;

    // cell merge

    /// <summary>
    /// Merges the cells within the range represented by the given string.
    /// </summary>
    void merge_cells(const std::string &reference_string);

    /// <summary>
    /// Merges the cells within the given range.
    /// </summary>
    void merge_cells(const range_reference &reference);

    /// <summary>
    /// Removes the merging of the cells in the range represented by the given string.
    /// </summary>
    void unmerge_cells(const std::string &reference_string);

    /// <summary>
    /// Removes the merging of the cells in the given range.
    /// </summary>
    void unmerge_cells(const range_reference &reference);

    /// <summary>
    /// Returns a vector of references of all merged ranges in the worksheet.
    /// </summary>
    std::vector<range_reference> merged_ranges() const;

    // operators

    /// <summary>
    /// Returns true if this worksheet refers to the same worksheet as other.
    /// </summary>
    bool operator==(const worksheet &other) const;

    /// <summary>
    /// Returns true if this worksheet doesn't refer to the same worksheet as other.
    /// </summary>
    bool operator!=(const worksheet &other) const;

    /// <summary>
    /// Returns true if this worksheet is null.
    /// </summary>
    bool operator==(std::nullptr_t) const;

    /// <summary>
    /// Returns true if this worksheet is not null.
    /// </summary>
    bool operator!=(std::nullptr_t) const;

    /// <summary>
    /// Sets the internal pointer of this worksheet object to point to other.
    /// </summary>
    void operator=(const worksheet &other);

    /// <summary>
    /// Convenience method for worksheet::cell method.
    /// </summary>
    class cell operator[](const cell_reference &reference);

    /// <summary>
    /// Convenience method for worksheet::cell method.
    /// </summary>
    const class cell operator[](const cell_reference &reference) const;

    /// <summary>
    /// Returns true if this worksheet is equal to other. If reference is true, the comparison
    /// will only check that both worksheets point to the same sheet in the same workbook.
    /// </summary>
    bool compare(const worksheet &other, bool reference) const;

    // page

    /// <summary>
    /// Returns true if this worksheet has a page setup.
    /// </summary>
    bool has_page_setup() const;

    /// <summary>
    /// Returns the page setup for this sheet.
    /// </summary>
    xlnt::page_setup page_setup() const;

    /// <summary>
    /// Sets the page setup for this sheet to setup.
    /// </summary>
    void page_setup(const struct page_setup &setup);

    /// <summary>
    /// Returns true if this page has margins.
    /// </summary>
    bool has_page_margins() const;

    /// <summary>
    /// Returns the margins of this sheet.
    /// </summary>
    xlnt::page_margins page_margins() const;

    /// <summary>
    /// Sets the margins of this sheet to margins.
    /// </summary>
    void page_margins(const class page_margins &margins);

    // auto filter

    /// <summary>
    /// Returns the current auto-filter of this sheet.
    /// </summary>
    range_reference auto_filter() const;

    /// <summary>
    /// Sets the auto-filter of this sheet to the range defined by range_string.
    /// </summary>
    void auto_filter(const std::string &range_string);

    /// <summary>
    /// Sets the auto-filter of this sheet to the given range.
    /// </summary>
    void auto_filter(const xlnt::range &range);

    /// <summary>
    /// Sets the auto-filter of this sheet to the range defined by reference.
    /// </summary>
    void auto_filter(const range_reference &reference);

    /// <summary>
    /// Clear the auto-filter from this sheet.
    /// </summary>
    void clear_auto_filter();

    /// <summary>
    /// Returns true if this sheet has an auto-filter set.
    /// </summary>
    bool has_auto_filter() const;

    /// <summary>
    /// Reserve n rows. This can be optionally called before adding many rows
    /// to improve performance.
    /// </summary>
    void reserve(std::size_t n);

    /// <summary>
    /// Returns true if this sheet has phonetic properties
    /// </summary>
    bool has_phonetic_properties() const;

    /// <summary>
    /// Returns the phonetic properties of this sheet.
    /// </summary>
    const phonetic_pr &phonetic_properties() const;

    /// <summary>
    /// Sets the phonetic properties of this sheet to phonetic_props
    /// </summary>
    void phonetic_properties(const phonetic_pr &phonetic_props);

    /// <summary>
    /// Returns true if this sheet has a header/footer.
    /// </summary>
    bool has_header_footer() const;

    /// <summary>
    /// Returns the header/footer of this sheet.
    /// </summary>
    class header_footer header_footer() const;

    /// <summary>
    /// Sets the header/footer of this sheet to new_header_footer.
    /// </summary>
    void header_footer(const class header_footer &new_header_footer);

    /// <summary>
    /// Returns the sheet state of this sheet.
    /// </summary>
    xlnt::sheet_state sheet_state() const;

    /// <summary>
    /// Sets the sheet state of this sheet.
    /// </summary>
    void sheet_state(xlnt::sheet_state state);

    /// <summary>
    /// Returns an iterator to the first row in this sheet.
    /// </summary>
    iterator begin();

    /// <summary>
    /// Returns an iterator one past the last row in this sheet.
    /// </summary>
    iterator end();

    /// <summary>
    /// Return a constant iterator to the first row in this sheet.
    /// </summary>
    const_iterator begin() const;

    /// <summary>
    /// Returns a constant iterator to one past the last row in this sheet.
    /// </summary>
    const_iterator end() const;

    /// <summary>
    /// Return a constant iterator to the first row in this sheet.
    /// </summary>
    const_iterator cbegin() const;

    /// <summary>
    /// Returns a constant iterator to one past the last row in this sheet.
    /// </summary>
    const_iterator cend() const;

    /// <summary>
    /// Sets rows to repeat at top during printing.
    /// </summary>
    void print_title_rows(row_t first_row, row_t last_row);

    /// <summary>
    /// Sets rows to repeat at top during printing.
    /// </summary>
    void print_title_rows(row_t last_row);

    /// <summary>
    /// Sets columns to repeat at left during printing.
    /// </summary>
    void print_title_cols(column_t first_column, column_t last_column);

    /// <summary>
    /// Sets columns to repeat at left during printing.
    /// </summary>
    void print_title_cols(column_t last_column);

    /// <summary>
    /// Returns a string representation of the defined print titles.
    /// </summary>
    std::string print_titles() const;

    /// <summary>
    /// Sets the print area of this sheet to print_area.
    /// </summary>
    void print_area(const std::string &print_area);

    /// <summary>
    /// Returns the print area defined for this sheet.
    /// </summary>
    range_reference print_area() const;

    /// <summary>
    /// Returns true if this sheet has any number of views defined.
    /// </summary>
    bool has_view() const;

    /// <summary>
    /// Returns the view at the given index.
    /// </summary>
    sheet_view &view(std::size_t index = 0) const;

    /// <summary>
    /// Adds new_view to the set of available views for this sheet.
    /// </summary>
    void add_view(const sheet_view &new_view);

    /// <summary>
    /// Set the active cell on the default worksheet view to the given reference.
    /// </summary>
    void active_cell(const cell_reference &ref);

    /// <summary>
    /// Returns true if the worksheet has a view and the view has an active cell.
    /// </summary>
    bool has_active_cell() const;

    /// <summary>
    /// Returns the active cell on the default worksheet view.
    /// </summary>
    cell_reference active_cell() const;

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

    /// <summary>
    /// Creates a conditional format for the given range with the given condition.
    /// </summary>
    xlnt::conditional_format conditional_format(const range_reference &ref, const condition &when);

    /// <summary>
    /// Returns the path of this worksheet in the containing package.
    /// </summary>
    xlnt::path path() const;

    /// <summary>
    /// Returns the relationship from the parent workbook to this worksheet.
    /// </summary>
    relationship referring_relationship() const;

    /// <summary>
    /// Returns the current formatting properties.
    /// </summary>
    sheet_format_properties format_properties() const;

    /// <summary>
    /// Sets the format properties to the given properties.
    /// </summary>
    void format_properties(const sheet_format_properties &properties);

    /// <summary>
    /// Returns true if this worksheet has a page setup.
    /// </summary>
    bool has_drawing() const;

    /// <summary>
    /// Returns true if this worksheet is empty.
    /// A worksheet is considered empty if it doesn't have any cells.
    /// </summary>
    bool is_empty() const;

private:
    friend class cell;
    friend class const_range_iterator;
    friend class range_iterator;
    friend class workbook;
    friend class detail::xlsx_consumer;
    friend class detail::xlsx_producer;

    /// <summary>
    /// Constructs a worksheet impl wrapper from d.
    /// </summary>
    worksheet(detail::worksheet_impl *d);

    /// <summary>
    /// Creates a comments part in the manifest as a relationship target of this sheet.
    /// </summary>
    void register_comments_in_manifest();

    /// <summary>
    /// Creates a calcChain part in the manifest if it doesn't already exist.
    /// </summary>
    void register_calc_chain_in_manifest();

    /// <summary>
    /// Removes calcChain part from manifest if no formulae remain in workbook.
    /// </summary>
    void garbage_collect_formulae();

    /// <summary>
    /// Sets the parent of this worksheet to wb.
    /// </summary>
    void parent(class workbook &wb);

    /// <summary>
    /// Move cells after index down or right by a given amount. The direction is decided by row_or_col.
    /// If reverse is true, the cells will be moved up or left, depending on row_or_col.
    /// </summary>
    void move_cells(std::uint32_t index, std::uint32_t amount, row_or_col_t row_or_col, bool reverse = false);

    /// <summary>
    /// The pointer to this sheet's implementation.
    /// </summary>
    detail::worksheet_impl *d_;
};

} // namespace xlnt
