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

#include <memory>
#include <string>
#include <unordered_map>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/cell/cell_type.hpp>
#include <xlnt/cell/index_types.hpp>

namespace xlnt {

enum class calendar;

class alignment;
class base_format;
class border;
class cell_reference;
class comment;
class fill;
class font;
class format;
class number_format;
class protection;
class style;
class workbook;
class worksheet;
class xlsx_consumer;
class xlsx_producer;

struct date;
struct datetime;
struct time;
struct timedelta;

namespace detail {

class xlsx_consumer;
class xlsx_producer;

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
class XLNT_API cell
{
public:
    /// <summary>
    /// Alias xlnt::cell_type to xlnt::cell::type since it looks nicer.
    /// </summary>
    using type = cell_type;

    /// <summary>
    /// Return a map of error strings such as \#DIV/0! and their associated indices.
    /// </summary>
    static const std::unordered_map<std::string, int> &error_codes();

    /// <summary>
    /// Default copy constructor.
    /// </summary>
    cell(const cell &) = default;

    // value

    /// <summary>
    /// Return true if value has been set and has not been cleared using cell::clear_value().
    /// </summary>
    bool has_value() const;

    /// <summary>
    /// Return the value of this cell as an instance of type T.
    /// Overloads exist for most C++ fundamental types like bool, int, etc. as well
    /// as for std::string and xlnt datetime types: date, time, datetime, and timedelta.
    /// </summary>
    template <typename T>
    T value() const;

    /// <summary>
    /// Make this cell have a value of type null.
    /// All other cell attributes are retained.
    /// </summary>
    void clear_value();

    /// <summary>
    /// Set the value of this cell to the given value.
    /// Overloads exist for most C++ fundamental types like bool, int, etc. as well
    /// as for std::string and xlnt datetime types: date, time, datetime, and timedelta.
    /// </summary>
    template <typename T>
    void value(T value);

    /// <summary>
    /// Analyze string_value to determine its type, convert it to that type,
    /// and set the value of this cell to that converted value.
    /// </summary>
    void value(const std::string &string_value, bool infer_type);

    /// <summary>
    /// Return the type of this cell.
    /// </summary>
    type data_type() const;

    /// <summary>
    /// Set the type of this cell.
    /// </summary>
    void data_type(type t);

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

    /// <summary>
    /// Return a cell_reference that points to the location of this cell.
    /// </summary>
    cell_reference reference() const;

    /// <summary>
    /// Return the column of this cell.
    /// </summary>
    column_t column() const;

    /// <summary>
    /// Return the row of this cell.
    /// </summary>
    row_t row() const;

    /// <summary>
    /// Return the location of this cell as an ordered pair.
    /// </summary>
    std::pair<int, int> anchor() const;

    // hyperlink

    /// <summary>
    /// Return the URL of this cell's hyperlink.
    /// </summary>
    std::string hyperlink() const;

    /// <summary>
    /// Add a hyperlink to this cell pointing to the URI of the given value.
    /// </summary>
    void hyperlink(const std::string &value);

    /// <summary>
    /// Return true if this cell has a hyperlink set.
    /// </summary>
    bool has_hyperlink() const;

    // computed formatting

    /// <summary>
    /// Returns the result of computed_format().alignment().
    /// </summary>
    class alignment computed_alignment() const;

    /// <summary>
    /// Returns the result of computed_format().border().
    /// </summary>
    class border computed_border() const;

    /// <summary>
    /// Returns the result of computed_format().fill().
    /// </summary>
    class fill computed_fill() const;

    /// <summary>
    /// Returns the result of computed_format().font().
    /// </summary>
    class font computed_font() const;

    /// <summary>
    /// Returns the result of computed_format().number_format().
    /// </summary>
    class number_format computed_number_format() const;

    /// <summary>
    /// Returns the result of computed_format().protection().
    /// </summary>
    class protection computed_protection() const;

    // format

    /// <summary>
    /// Return true if this cell has had a format applied to it.
    /// </summary>
    bool has_format() const;

    /// <summary>
    /// Return a reference to the format applied to this cell.
    /// If this cell has no format, an invalid_attribute exception will be thrown.
    /// </summary>
    const class format format() const;

    /// <summary>
    /// Applies the cell-level formatting of new_format to this cell.
    /// </summary>
    void format(const class format new_format);

    /// <summary>
    /// Remove the cell-level formatting from this cell.
    /// This doesn't affect the style that may also be applied to the cell.
    /// Throws an invalid_attribute exception if no format is applied.
    /// </summary>
    void clear_format();

    /// <summary>
    /// Returns the number format of this cell.
    /// </summary>
    class number_format number_format() const;

    /// <summary>
    /// Creates a new format in the workbook, sets its number_format
    /// to the given format, and applies the format to this cell.
    /// </summary>
    void number_format(const class number_format &format);

    /// <summary>
    /// Returns the font applied to the text in this cell.
    /// </summary>
    class font font() const;

    /// <summary>
    /// Creates a new format in the workbook, sets its font
    /// to the given font, and applies the format to this cell.
    /// </summary>
    void font(const class font &font_);

    /// <summary>
    /// Returns the fill applied to this cell.
    /// </summary>
    class fill fill() const;

    /// <summary>
    /// Creates a new format in the workbook, sets its fill
    /// to the given fill, and applies the format to this cell.
    /// </summary>
    void fill(const class fill &fill_);

    /// <summary>
    /// Returns the border of this cell.
    /// </summary>
    class border border() const;

    /// <summary>
    /// Creates a new format in the workbook, sets its border
    /// to the given border, and applies the format to this cell.
    /// </summary>
    void border(const class border &border_);

    /// <summary>
    /// Returns the alignment of the text in this cell.
    /// </summary>
    class alignment alignment() const;

    /// <summary>
    /// Creates a new format in the workbook, sets its alignment
    /// to the given alignment, and applies the format to this cell.
    /// </summary>
    void alignment(const class alignment &alignment_);

    /// <summary>
    /// Returns the protection of this cell.
    /// </summary>
    class protection protection() const;

    /// <summary>
    /// Creates a new format in the workbook, sets its protection
    /// to the given protection, and applies the format to this cell.
    /// </summary>
    void protection(const class protection &protection_);

    // style

    /// <summary>
    /// Returns true if this cell has had a style applied to it.
    /// </summary>
    bool has_style() const;

    /// <summary>
    /// Returns a wrapper pointing to the named style applied to this cell.
    /// </summary>
    const class style style() const;

    /// <summary>
    /// Equivalent to style(new_style.name())
    /// </summary>
    void style(const class style &new_style);

    /// <summary>
    /// Sets the named style applied to this cell to a style named style_name.
    /// If this style has not been previously created in the workbook, a
    /// key_not_found exception will be thrown.
    /// </summary>
    void style(const std::string &style_name);

    /// <summary>
    /// Removes the named style from this cell.
    /// An invalid_attribute exception will be thrown if this cell has no style.
    /// This will not affect the cell format of the cell.
    /// </summary>
    void clear_style();

    // formula

    /// <summary>
    /// Returns the string representation of the formula applied to this cell.
    /// </summary>
    std::string formula() const;

    /// <summary>
    /// Sets the formula of this cell to the given value.
    /// This formula string should begin with '='.
    /// </summary>
    void formula(const std::string &formula);

    /// <summary>
    /// Removes the formula from this cell. After this is called, has_formula() will return false.
    /// </summary>
    void clear_formula();

    /// <summary>
    /// Returns true if this cell has had a formula applied to it.
    /// </summary>
    bool has_formula() const;

    // printing

    /// <summary>
    /// Returns a string representing the value of this cell. If the data type is not a string,
    /// it will be converted according to the number format.
    /// </summary>
    std::string to_string() const;

    // merging

    /// <summary>
    /// Return true iff this cell has been merged with one or more
    /// surrounding cells.
    /// </summary>
    bool is_merged() const;

    /// <summary>
    /// Make this a merged cell iff merged is true.
    /// Generally, this shouldn't be called directly. Instead,
    /// use worksheet::merge_cells on its parent worksheet.
    /// </summary>
    void merged(bool merged);

    /// <summary>
    /// Return the error string that is stored in this cell.
    /// </summary>
    std::string error() const;

    /// <summary>
    /// Directly assign the value of this cell to be the given error.
    /// </summary>
    void error(const std::string &error);

    /// <summary>
    /// Return a cell from this cell's parent workbook at
    /// a relative offset given by the parameters.
    /// </summary>
    cell offset(int column, int row);

    /// <summary>
    /// Return the worksheet that owns this cell.
    /// </summary>
    class worksheet worksheet();

    /// <summary>
    /// Return the worksheet that owns this cell.
    /// </summary>
    const class worksheet worksheet() const;

    /// <summary>
    /// Return the workbook of the worksheet that owns this cell.
    /// </summary>
    class workbook &workbook();

    /// <summary>
    /// Return the workbook of the worksheet that owns this cell.
    /// </summary>
    const class workbook &workbook() const;

    /// <summary>
    /// Shortcut to return the base date of the parent workbook.
    /// Equivalent to workbook().properties().excel_base_date
    /// </summary>
    calendar base_date() const;

    /// <summary>
    /// Return to_check after checking encoding, size, and illegal characters.
    /// </summary>
    std::string check_string(const std::string &to_check);

    // comment

    /// <summary>
    /// Return true if this cell has a comment applied.
    /// </summary>
    bool has_comment();

    /// <summary>
    /// Delete the comment applied to this cell if it exists.
    /// </summary>
    void clear_comment();

    /// <summary>
    /// Get the comment applied to this cell.
    /// </summary>
    class comment comment();

    /// <summary>
    /// Create a new comment with the given text and optional author and apply it to the cell.
    /// </summary>
    void comment(const std::string &text, const std::string &author = "Microsoft Office User");

    /// <summary>
    /// Apply the comment provided as the only argument to the cell.
    /// </summary>
    void comment(const class comment &new_comment);

    // operators

    /// <summary>
    /// Make this cell point to rhs.
    /// The cell originally pointed to by this cell will be unchanged.
    /// </summary>
    cell &operator=(const cell &rhs);

    /// <summary>
    /// Return true if this cell the same cell as comparand (compare by reference).
    /// </summary>
    bool operator==(const cell &comparand) const;

    /// <summary>
    /// Return true if this cell is uninitialized.
    /// </summary>
    bool operator==(std::nullptr_t) const;

private:
    // make these friends so they can use the private constructor
    friend class style;
    friend class worksheet;
    friend class detail::xlsx_consumer;
    friend class detail::xlsx_producer;
    friend struct detail::cell_impl;

    /// <summary>
    /// Returns a non-const reference to the format of this cell.
    /// This is for internal use only.
    /// </summary>
    class format modifiable_format();

    /// <summary>
    /// Delete default zero-argument constructor.
    /// </summary>
    cell() = delete;

    /// <summary>
    /// Private constructor to create a cell from its implementation.
    /// </summary>
    cell(detail::cell_impl *d);

    /// <summary>
    /// A pointer to this cell's implementation.
    /// </summary>
    detail::cell_impl *d_;
};

/// <summary>
/// Return true if this cell is uninitialized.
/// </summary>
XLNT_API bool operator==(std::nullptr_t, const cell &cell);

/// <summary>
/// Convenience function for writing cell to an ostream.
/// Uses cell::to_string() internally.
/// </summary>
XLNT_API std::ostream &operator<<(std::ostream &stream, const xlnt::cell &cell);

} // namespace xlnt
