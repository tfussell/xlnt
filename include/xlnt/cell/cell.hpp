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

#include <xlnt/xlnt_config.hpp> // for XLNT_CLASS, XLNT_FUNCTION
#include <xlnt/cell/cell_type.hpp> // for cell_type
#include <xlnt/cell/index_types.hpp> // for column_t, row_t

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

namespace detail { struct cell_impl; }

/// <summary>
/// Describes cell associated properties.
/// </summary>
/// <remarks>
/// Properties of interest include style, type, value, and address.
/// The Cell class is required to know its value and type, display options,
/// and any other features of an Excel cell.Utilities for referencing
/// cells using Excel's 'A1' column/row nomenclature are also provided.
/// </remarks>
class XLNT_CLASS cell
{
public:
    using type = cell_type;
    
    /// <summary>
    /// Return a map of error strings such as \#DIV/0! and their associated indices.
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

    /// <summary>
    /// Return the value of this cell as an instance of type T.
    /// Overloads exist for most C++ fundamental types like bool, int, etc. as well
    /// as for std::string and xlnt datetime types: date, time, datetime, and timedelta.
    /// </summary>
    template <typename T>
    T get_value() const;

    /// <summary>
    /// Make this cell have a value of type null.
    /// All other cell attributes are retained.
    /// </summary>
    void clear_value();

    /// <summary>
    /// Set the value of this cell to the given value.
    /// Overloads exist for most C++ fundamental types like bool, int, etc. as well
    /// as for std::string and xlnt datetime types: date, time, datetime, and timedelta.
    template <typename T>
    void set_value(T value);

    /// <summary>
    /// Return the type of this cell.
    /// </summary>
    type get_data_type() const;
    
    /// <summary>
    /// Set the type of this cell.
    /// </summary>
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
    
    /// <summary>
    /// Return a cell_reference that points to the location of this cell.
    /// </summary>
    cell_reference get_reference() const;
    
    /// <summary>
    /// Return the column of this cell.
    /// </summary>
    column_t get_column() const;
    
    /// <summary>
    /// Return the row of this cell.
    /// </summary>
    row_t get_row() const;
    
    /// <summary>
    /// Return the location of this cell as an ordered pair.
    /// </summary>
    std::pair<int, int> get_anchor() const;

    // hyperlink
    
    /// <summary>
    /// Return a relationship representing this cell's hyperlink.
    /// </summary>
    relationship get_hyperlink() const;
    
    /// <summary>
    /// Add a hyperlink to this cell pointing to the URI of the given value.
    /// </summary>
    void set_hyperlink(const std::string &value);
    
    /// <summary>
    /// Return true if this cell has a hyperlink set.
    /// </summary>
    bool has_hyperlink() const;

    // style
    
    /// <summary>
    /// Return true if this cell has had a style applied to it.
    /// </summary>
    bool has_style() const;
    
    /// <summary>
    /// Return the index of this cell's style in its parent workbook.
    /// This is also the index of the style in the stylesheet XML, xl/styles.xml.
    /// </summary>
    std::size_t get_style_id() const;
    
    /// <summary>
    /// Set the style index of this cell. This should be an existing style in
    /// the parent workbook.
    /// </summary>
    void set_style_id(std::size_t style_id);
    
    /// <summary>
    /// Return the number format of this cell.
    /// </summary>
    const number_format &get_number_format() const;
    void set_number_format(const number_format &format);
    
    /// <summary>
    /// Return the font applied to the text in this cell.
    /// </summary>
    const font &get_font() const;
    
    void set_font(const font &font_);
    
    /// <summary>
    /// Return the fill applied to this cell.
    /// </summary>
    const fill &get_fill() const;
    
    void set_fill(const fill &fill_);
    
    /// <summary>
    /// Return the border of this cell.
    /// </summary>
    const border &get_border() const;
    
    void set_border(const border &border_);
    
    /// <summary>
    /// Return the alignment of the text in this cell.
    /// </summary>
    const alignment &get_alignment() const;
    
    void set_alignment(const alignment &alignment_);
    
    /// <summary>
    /// Return the protection of this cell.
    /// </summary>
    const protection &get_protection() const;
    
    void set_protection(const protection &protection_);
    
    void set_pivot_button(bool b);
    
    /// <summary>
    /// Return true iff pivot button?
    /// </summary>
    bool pivot_button() const;
    
    void set_quote_prefix(bool b);
    
    /// <summary>
    /// Return true iff quote prefix?
    /// </summary>
    bool quote_prefix() const;

    // comment
    
    /// <summary>
    /// Return the comment of this cell.
    /// </summary>
    comment get_comment();
    
    /// <summary>
    /// Return the comment of this cell.
    /// </summary>
    const comment get_comment() const;
    
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
    void set_merged(bool merged);

    /// <summary>
    /// Return the error string that is stored in this cell.
    /// </summary>
    std::string get_error() const;
    
    /// <summary>
    /// Directly assign the value of this cell to be the given error.
    /// </summary>
    void set_error(const std::string &error);

    /// <summary>
    /// Return a cell from this cell's parent workbook at
    /// a relative offset given by the parameters.
    /// </summary>
    cell offset(int column, int row);
    
    /// <summary>
    /// Return the worksheet that owns this cell.
    /// </summary>
    worksheet get_parent();
    
    /// <summary>
    /// Return the worksheet that owns this cell.
    /// </summary>
    const worksheet get_parent() const;

    /// <summary>
    /// Shortcut to return the base date of the parent workbook.
    /// Equivalent to get_parent().get_parent().get_properties().excel_base_date
    /// </summary>
    calendar get_base_date() const;

    /// <summary>
    /// Return to_check after checking encoding, size, and illegal characters.
    /// </summary>
    std::string check_string(const std::string &to_check);

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

    // friend operators, so we can put cell on either side of comparisons with other types
    
    /// <summary>
    /// Return true if this cell is uninitialized.
    /// </summary>
    friend XLNT_FUNCTION bool operator==(std::nullptr_t, const cell &cell);
    
    /// <summary>
    /// Return the result of left.get_reference() < right.get_reference().
    /// What's the point of this?
    /// </summary>
    friend XLNT_FUNCTION bool operator<(cell left, cell right);

	/// <summary>
	/// Convenience function for writing cell to an ostream.
	/// Uses cell::to_string() internally.
	/// </summary>
	friend XLNT_FUNCTION std::ostream &operator<<(std::ostream &stream, const xlnt::cell &cell);

private:
    // make these friends so they can use the private constructor
    friend class style;
    friend class worksheet;
    friend struct detail::cell_impl;

    /// <summary>
    /// Private constructor to create a cell from its implementation.
    /// </summary>
    cell(detail::cell_impl *d);
    
    /// <summary>
    /// A pointer to this cell's implementation.
    /// </summary>
    detail::cell_impl *d_;
};

} // namespace xlnt
