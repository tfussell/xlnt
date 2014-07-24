// Copyright (c) 2014 Thomas Fussell
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
    
    static std::pair<std::string, row_t> split_reference(const std::string &reference_string,
        bool &absolute_column, bool &absolute_row);
    
    cell_reference();
    cell_reference(const char *reference_string);
    cell_reference(const std::string &reference_string);
    cell_reference(const std::string &column, row_t row, bool absolute = false);
    cell_reference(column_t column, row_t row, bool absolute = false);
    
    bool is_absolute() const { return absolute_; }
    void set_absolute(bool absolute) { absolute_ = absolute; }
    
    std::string get_column() const { return column_string_from_index(column_index_ + 1); }
    void set_column(const std::string &column) { column_index_ = column_index_from_string(column) - 1; }
    
    column_t get_column_index() const { return column_index_; }
    void set_column_index(column_t column_index) { column_index_ = column_index; }
    
    row_t get_row() const { return row_index_ + 1; }
    void set_row(row_t row) { row_index_ = row - 1; }
    
    row_t get_row_index() const { return row_index_; }
    void set_row_index(row_t row_index) { row_index_ = row_index; }
    
    cell_reference make_offset(int column_offset, int row_offset) const;
    
    std::string to_string() const;
    range_reference to_range() const;
    
    bool operator==(const cell_reference &comparand) const;
    bool operator==(const std::string &reference_string) const { return *this == cell_reference(reference_string); }
    bool operator==(const char *reference_string) const { return *this == std::string(reference_string); }
    bool operator!=(const cell_reference &comparand) const { return !(*this == comparand); }
    bool operator!=(const std::string &reference_string) const { return *this != cell_reference(reference_string); }
    bool operator!=(const char *reference_string) const { return *this != std::string(reference_string); }

    friend bool operator<(const cell_reference &left, const cell_reference &right);

private:
    column_t column_index_;
    row_t row_index_;
    bool absolute_;
};
   
} // namespace xlnt
