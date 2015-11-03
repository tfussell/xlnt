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

#include <memory>
#include <string>
#include <unordered_map>

#include <xlnt/cell/types.hpp>

namespace xlnt {

enum class calendar;

class alignment;
class border;
class cell_reference;
class comment;
class fill;
class font;
class number_format;
class protection;
class relationship;
class worksheet;

struct date;
struct datetime;
struct time;
struct timedelta;

namespace detail {

struct cell_impl;

} // namespace detail

/// <summary>
/// Describes cell associated properties.
/// </summary>
/// <remarks>
/// Properties of interest include style, type, value, and address.
/// The Cell class is required to know its value and type, display options,
/// and any other features of an Excel cell.Utilities for referencing
/// cells using Excel's 'A1' column/row nomenclature are also provided.
/// </remarks>
class cell
{
  public:
    /// <summary>
    /// Enumerates the possible types a cell can be determined by it's current value.
    /// </summary>
    enum class type
    {
        null,
        numeric,
        string,
        formula,
        error,
        boolean
    };

    /// <summary>
    /// Return a map of error strings such as #DIV/0! and their associated indices.
    /// </summary>
    static const std::unordered_map<std::string, int> &error_codes();

    // TODO: Should it be possible to construct and use a cell without a parent worksheet?
    //(cont'd) If so, it would need to be responsible for allocating and deleting its PIMPL.

    /// <summary>
    /// Construct a null cell without a parent.
    /// Most methods will throw an exception if this cell is not further initialized.
    /// </summary>
    cell();

    /// <summary>
    /// Construct a cell in worksheet, sheet, at the given reference location (e.g. A1).
    /// </summary>
    cell(worksheet sheet, const cell_reference &reference);

    /// <summary>
    /// This constructor, provided for convenience, is equivalent to calling:
    /// cell c(sheet, reference);
    /// c.set_value(initial_value);
    /// </summary>
    template <typename T>
    cell(worksheet sheet, const cell_reference &reference, const T &initial_value);

    // value

    /// <summary>
    /// Return true if value has been set and has not been cleared using cell::clear_value().
    /// </summary>
    bool has_value() const;

    template <typename T>
    T get_value() const;

    void clear_value();

    template <typename T>
    void set_value(T value);

    type get_data_type() const;
    void set_data_type(type t);

    // properties

    /// <summary>
    /// There's no reason to keep a cell which has no value and is not a placeholder.
    /// Return true if this cell has no value, style, isn't merged, etc.
    /// </summary>
    bool garbage_collectible() const;

    /// <summary>
    /// Return true iff this cell's number format matches a date format.
    /// </summary>
    bool is_date() const;

    // position
    cell_reference get_reference() const;
    std::string get_column() const;
    column_t get_column_index() const;
    row_t get_row() const;
    std::pair<int, int> get_anchor() const;

    // hyperlink
    relationship get_hyperlink() const;
    void set_hyperlink(const std::string &value);
    bool has_hyperlink() const;

    // style
    bool has_style() const;
    std::size_t get_style_id() const;
    void set_style_id(std::size_t style_id);
    const number_format &get_number_format() const;
    void set_number_format(const number_format &format);
    const font &get_font() const;
    void set_font(const font &font_);
    const fill &get_fill() const;
    void set_fill(const fill &fill_);
    const border &get_border() const;
    void set_border(const border &border_);
    const alignment &get_alignment() const;
    void set_alignment(const alignment &alignment_);
    const protection &get_protection() const;
    void set_protection(const protection &protection_);
    void set_pivot_button(bool b);
    bool pivot_button() const;
    void set_quote_prefix(bool b);
    bool quote_prefix() const;

    // comment
    comment get_comment();
    void set_comment(const comment &comment);
    void clear_comment();
    bool has_comment() const;

    // formula
    std::string get_formula() const;
    void set_formula(const std::string &formula);
    void clear_formula();
    bool has_formula() const;

    // printing

    /// <summary>
    /// Returns a string describing this cell like <Cell Sheet.A1>.
    /// </summary>
    std::string to_repr() const;

    /// <summary>
    /// Returns a string representing the value of this cell. If the data type is not a string,
    /// it will be converted according to the number format.
    /// </summary>
    std::string to_string() const;

    // merging
    bool is_merged() const;
    void set_merged(bool merged);

    std::string get_error() const;
    void set_error(const std::string &error);

    cell offset(row_t row, column_t column);

    worksheet get_parent();
    const worksheet get_parent() const;

    calendar get_base_date() const;

    // operators
    cell &operator=(const cell &rhs);

    bool operator==(const cell &comparand) const;
    bool operator==(std::nullptr_t) const;

    // friend operators, so we can put cell on either side of comparisons with other types
    friend bool operator==(std::nullptr_t, const cell &cell);
    friend bool operator<(cell left, cell right);

  private:
    friend class worksheet;
    friend struct detail::cell_impl;
    friend class style;

    cell(detail::cell_impl *d);
    detail::cell_impl *d_;
};

/// <summary>
/// Convenience function for writing cell to an ostream.
/// Uses cell::to_string() internally.
/// </summary>
inline std::ostream &operator<<(std::ostream &stream, const xlnt::cell &cell)
{
    return stream << cell.to_string();
}

} // namespace xlnt
