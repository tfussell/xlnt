// Copyright (c) 2015 Thomas Fussell
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

#include <cstdint>
#include <string>
#include <utility>

#include "../common/types.hpp"

namespace xlnt {
    
class cell_reference;
class range_reference;
    
struct cell_reference_hash
{
    std::size_t operator()(const cell_reference &k) const;
};
    
class cell_reference
{
public:
    /// <summary>
    /// Convert a coordinate to an absolute coordinate string (B12 -> $B$12)
    /// </summary>
    static cell_reference make_absolute(const cell_reference &relative_reference);
    
    /// <summary>
    /// Convert a column letter into a column number (e.g. B -> 2)
    /// </summary>
    /// <remarks>
    /// Excel only supports 1 - 3 letter column names from A->ZZZ, so we
    /// restrict our column names to 1 - 3 characters, each in the range A - Z.
    /// </remarks>
    static column_t column_index_from_string(const std::string &column_string);
    
    /// <summary>
    /// Convert a column number into a column letter (3 -> 'C')
    /// </summary>
    /// <remarks>
    /// Right shift the column col_idx by 26 to find column letters in reverse
    /// order.These numbers are 1 - based, and can be converted to ASCII
    /// ordinals by adding 64.
    /// </remarks>
    static std::string column_string_from_index(column_t column_index);
    
    /// <summary>
    /// Split a coordinate string like "A1" into an equivalent pair like {"A", 1}.
    /// </summary>
    static std::pair<std::string, row_t> split_reference(const std::string &reference_string,
        bool &absolute_column, bool &absolute_row);
    
    // constructors
    /// <summary>
    /// Default constructor makes a reference to the top-left-most cell, "A1".
    /// </summary>
    cell_reference();
    cell_reference(const char *reference_string);
    cell_reference(const std::string &reference_string);
    cell_reference(const std::string &column, row_t row, bool absolute = false);
    cell_reference(column_t column, row_t row, bool absolute = false);
    
    // absoluateness
    bool is_absolute() const { return absolute_; }
    void set_absolute(bool absolute) { absolute_ = absolute; }
    
    // getters/setters
    std::string get_column() const { return column_string_from_index(column_); }
    void set_column(const std::string &column_string) { column_ = column_index_from_string(column_string); }
    
    column_t get_column_index() const { return column_; }
    void set_column_index(column_t column) { column_ = column; }
    
    row_t get_row() const { return row_ ; }
    void set_row(row_t row) { row_ = row; }
    
    /// <summary>
    /// Return a cell_reference offset from this cell_reference by
    /// the number of columns and rows specified by the parameters.
    /// A negative value for column_offset or row_offset results
    /// in a reference above or left of this cell_reference, respectively.
    /// </summary>
    cell_reference make_offset(int column_offset, int row_offset) const;
    
    /// <summary>
    /// Return a string like "A1" for cell_reference(1, 1).
    /// </summary>
    std::string to_string() const;
    
    /// <summary>
    /// Return a range_reference containing only this cell_reference.
    /// </summary>
    range_reference to_range() const;
    
    // operators
    /// <summary>
    /// I've always wanted to overload the comma operator.
    /// cell_reference("A", 1), cell_reference("B", 1) will return
    /// range_reference(cell_reference("A", 1), cell_reference("B", 1))
    /// </summary>
    range_reference operator,(const cell_reference &other) const;
    
    bool operator==(const cell_reference &comparand) const;
    bool operator==(const std::string &reference_string) const { return *this == cell_reference(reference_string); }
    bool operator==(const char *reference_string) const { return *this == std::string(reference_string); }
    bool operator!=(const cell_reference &comparand) const { return !(*this == comparand); }
    bool operator!=(const std::string &reference_string) const { return *this != cell_reference(reference_string); }
    bool operator!=(const char *reference_string) const { return *this != std::string(reference_string); }

    bool operator<(const cell_reference &other);
    bool operator>(const cell_reference &other);
    bool operator<=(const cell_reference &other);
    bool operator>=(const cell_reference &other);

private:
    /// <summary>
    /// Index of the column. Important: this is one-indexed to conform
    /// with Excel. Column "A", the first column, would have an index of 1.
    /// </summary>
    column_t column_;
    
    /// <summary>
    /// Index of the column. Important: this is one-indexed to conform
    /// with Excel. Column "A", the first column, would have an index of 1.
    /// </summary>
    row_t row_;
    
    /// <summary>
    /// True if the reference is absolute. This looks like "$A$1" in Excel.
    /// </summary>
    bool absolute_;
};
   
} // namespace xlnt
