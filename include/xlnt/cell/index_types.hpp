// Copyright (c) 2014-2016 Thomas Fussell
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

#include <algorithm>
#include <cstdint>
#include <string>

#include <xlnt/xlnt_config.hpp>

// We might want to change these types for various optimizations in the future
// so use typedefs.

namespace xlnt {

/// <summary>
/// All rows should be referred to by an instance of this type.
/// </summary>
using row_t = std::uint32_t;

/// <summary>
/// Columns can be referred to as a string A,B,...Z,AA,AB,..,ZZ,AAA,...,ZZZ
/// or as a 1-indexed index. This class encapsulates both of these forms of
/// column referencing and allows for convertions between them.
/// </summary>
class XLNT_CLASS column_t
{
public:
    using index_t = std::uint32_t;

    /// <summary>
    /// Convert a column letter into a column number (e.g. B -> 2)
    /// </summary>
    /// <remarks>
    /// Excel only supports 1 - 3 letter column names from A->ZZZ, so we
    /// restrict our column names to 1 - 3 characters, each in the range A - Z.
    /// Strings outside this range and malformed strings will throw xlnt::column_string_index_exception.
    /// </remarks>
    static index_t column_index_from_string(const std::string &column_string);

    /// <summary>
    /// Convert a column number into a column letter (3 -> 'C')
    /// </summary>
    /// <remarks>
    /// Right shift the column, column_index, by 26 to find column letters in reverse
    /// order. These indices are 1-based, and can be converted to ASCII
    /// ordinals by adding 64.
    /// </remarks>
    static std::string column_string_from_index(index_t column_index);
    
    /// <summary>
    /// Default column_t is the first (left-most) column.
    /// </summary>
    column_t();
    
    /// <summary>
    /// Construct a column from a number.
    /// </summary>
    column_t(index_t column_index);
    
    /// <summary>
    /// Construct a column from a string.
    /// </summary>
    explicit column_t(const std::string &column_string);
    
    /// <summary>
    /// Construct a column from a string.
    /// </summary>
    explicit column_t(const char *column_string);
    
    /// <summary>
    /// Copy constructor
    /// </summary>
    column_t(const column_t &other);
    
    /// <summary>
    /// Move constructor
    /// </summary>
    column_t(column_t &&other);
    
    /// <summary>
    /// Return a string representation of this column index.
    /// </summary>
    std::string column_string() const;
    
    /// <summary>
    /// Set this column to be the same as rhs's and return reference to self.
    /// </summary>
    column_t &operator=(column_t rhs);
    
    /// <summary>
    /// Set this column to be equal to rhs and return reference to self.
    /// </summary>
    column_t &operator=(const std::string &rhs);
    
    /// <summary>
    /// Set this column to be equal to rhs and return reference to self.
    /// </summary>
    column_t &operator=(const char *rhs);
    
    /// <summary>
    /// Return true if this column refers to the same column as other.
    /// </summary>
    bool operator==(const column_t &other) const;
    
    /// <summary>
    /// Return true if this column doesn't refer to the same column as other.
    /// </summary>
    bool operator!=(const column_t &other) const;

    /// <summary>
    /// Return true if this column refers to the same column as other.
    /// </summary>
    bool operator==(int other) const;
    
    /// <summary>
    /// Return true if this column refers to the same column as other.
    /// </summary>
    bool operator==(index_t other) const;
    
    /// <summary>
    /// Return true if this column refers to the same column as other.
    /// </summary>
    bool operator==(const std::string &other) const;
    
    /// <summary>
    /// Return true if this column refers to the same column as other.
    /// </summary>
    bool operator==(const char *other) const;
    
    /// <summary>
    /// Return true if this column doesn't refer to the same column as other.
    /// </summary>
    bool operator!=(int other) const;
    
    /// <summary>
    /// Return true if this column doesn't refer to the same column as other.
    /// </summary>
    bool operator!=(index_t other) const;
    
    /// <summary>
    /// Return true if this column doesn't refer to the same column as other.
    /// </summary>
    bool operator!=(const std::string &other) const;
    
    /// <summary>
    /// Return true if this column doesn't refer to the same column as other.
    /// </summary>
    bool operator!=(const char *other) const;
    
    /// <summary>
    /// Return true if other is to the right of this column.
    /// </summary>
    bool operator>(const column_t &other) const;
    
    /// <summary>
    /// Return true if other is to the right of or equal to this column.
    /// </summary>
    bool operator>=(const column_t &other) const;
    
    /// <summary>
    /// Return true if other is to the left of this column.
    /// </summary>
    bool operator<(const column_t &other) const;
    
    /// <summary>
    /// Return true if other is to the left of or equal to this column.
    /// </summary>
    bool operator<=(const column_t &other) const;
    
    /// <summary>
    /// Return true if other is to the right of this column.
    /// </summary>
    bool operator>(const column_t::index_t &other) const;
    
    /// <summary>
    /// Return true if other is to the right of or equal to this column.
    /// </summary>
    bool operator>=(const column_t::index_t &other) const;
    
    /// <summary>
    /// Return true if other is to the left of this column.
    /// </summary>
    bool operator<(const column_t::index_t &other) const;
    
    /// <summary>
    /// Return true if other is to the left of or equal to this column.
    /// </summary>
    bool operator<=(const column_t::index_t &other) const;
    
    /// <summary>
    /// Pre-increment this column, making it point to the column one to the right.
    /// </summary>
    column_t &operator++();
    
    /// <summary>
    /// Pre-deccrement this column, making it point to the column one to the left.
    /// </summary>
    column_t &operator--();
    
    /// <summary>
    /// Post-increment this column, making it point to the column one to the right and returning the old column.
    /// </summary>
    column_t operator++(int);
    
    /// <summary>
    /// Post-decrement this column, making it point to the column one to the left and returning the old column.
    /// </summary>
    column_t operator--(int);
    
    /// <summary>
    /// Return the result of adding rhs to this column.
    /// </summary>
    column_t operator+(const column_t &rhs);
    
    /// <summary>
    /// Return the result of adding rhs to this column.
    /// </summary>
    column_t operator-(const column_t &rhs);
    
    /// <summary>
    /// Return the result of adding rhs to this column.
    /// </summary>
    column_t operator*(const column_t &rhs);
    
    /// <summary>
    /// Return the result of adding rhs to this column.
    /// </summary>
    column_t operator/(const column_t &rhs);
    
    /// <summary>
    /// Return the result of adding rhs to this column.
    /// </summary>
    column_t operator%(const column_t &rhs);
    
    /// <summary>
    /// Add rhs to this column and return a reference to this column.
    /// </summary>
    column_t &operator+=(const column_t &rhs);
    
    /// <summary>
    /// Subtrac rhs from this column and return a reference to this column.
    /// </summary>
    column_t &operator-=(const column_t &rhs);
    
    /// <summary>
    /// Multiply this column by rhs and return a reference to this column.
    /// </summary>
    column_t &operator*=(const column_t &rhs);
    
    /// <summary>
    /// Divide this column by rhs and return a reference to this column.
    /// </summary>
    column_t &operator/=(const column_t &rhs);
    
    /// <summary>
    /// Mod this column by rhs and return a reference to this column.
    /// </summary>
    column_t &operator%=(const column_t &rhs);
    
    /// <summary>
    /// Return true if other is to the right of this column.
    /// </summary>
    friend bool operator>(const column_t::index_t &left, const column_t &right);
    
    /// <summary>
    /// Return true if other is to the right of or equal to this column.
    /// </summary>
    friend bool operator>=(const column_t::index_t &left, const column_t &right);
    
    /// <summary>
    /// Return true if other is to the left of this column.
    /// </summary>
    friend bool operator<(const column_t::index_t &left, const column_t &right);
    
    /// <summary>
    /// Return true if other is to the left of or equal to this column.
    /// </summary>
    friend bool operator<=(const column_t::index_t &left, const column_t &right);
    
    /// <summary>
    /// Swap the columns that left and right refer to.
    /// </summary>
    friend void swap(column_t &left, column_t &right);
    
    /// <summary>
    /// Internal numeric value of this column index.
    /// </summary>
    index_t index;
};

/// <summary>
/// Functor for hashing a column.
/// Allows for use of std::unordered_set<column, column_hash> and similar.
/// </summary>
struct XLNT_CLASS column_hash
{
    std::size_t operator()(const column_t &k) const;
};

} // namespace xlnt

namespace std {

template <>
struct hash<xlnt::column_t>
{
    size_t operator()(const xlnt::column_t &k) const
    {
        return hasher(k);
    }
    
    xlnt::column_hash hasher;
};

} // namespace std
