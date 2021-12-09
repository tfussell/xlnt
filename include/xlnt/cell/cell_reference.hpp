// Copyright (c) 2014-2021 Thomas Fussell
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

#include <xlnt/xlnt_config.hpp>
#include <xlnt/cell/index_types.hpp>
#include <cstdint>
#include <string>
#include <tuple>
#include <utility>

namespace xlnt {

class cell_reference;
class range_reference;

/// <summary>
/// Functor for hashing a cell reference.
/// Allows for use of std::unordered_set<cell_reference, cel_reference_hash> and similar.
/// </summary>
struct XLNT_API cell_reference_hash
{
    /// <summary>
    /// Returns a hash representing a particular cell_reference.
    /// </summary>
    std::size_t operator()(const cell_reference &k) const;
};

/// <summary>
/// An object used to refer to a cell.
/// References have two parts, the column and the row.
/// In Excel, the reference string A1 refers to the top-left-most cell. A cell_reference
/// can be initialized from a string of this form or a 1-indexed ordered pair of the form
/// column, row.
/// </summary>
class XLNT_API cell_reference
{
public:
    /// <summary>
    /// Splits a coordinate string like "A1" into an equivalent pair like {"A", 1}.
    /// </summary>
    static std::pair<std::string, row_t> split_reference(const std::string &reference_string);

    /// <summary>
    /// Splits a coordinate string like "A1" into an equivalent pair like {"A", 1}.
    /// Reference parameters absolute_column and absolute_row will be set to true
    /// if column part or row part are prefixed by a dollar-sign indicating they
    /// are absolute, otherwise false.
    /// </summary>
    static std::pair<std::string, row_t> split_reference(
        const std::string &reference_string, bool &absolute_column, bool &absolute_row);

    // constructors

    /// <summary>
    /// Default constructor makes a reference to the top-left-most cell, "A1".
    /// </summary>
    cell_reference();

    // TODO: should these be explicit? The implicit conversion is nice sometimes.

    /// <summary>
    /// Constructs a cell_reference from a string reprenting a cell coordinate (e.g. $B14).
    /// </summary>
    cell_reference(const char *reference_string);

    /// <summary>
    /// Constructs a cell_reference from a string reprenting a cell coordinate (e.g. $B14).
    /// </summary>
    cell_reference(const std::string &reference_string);

    /// <summary>
    /// Constructs a cell_reference from a 1-indexed column index and row index.
    /// </summary>
    cell_reference(column_t column, row_t row);

    // absoluteness

    /// <summary>
    /// Converts a coordinate to an absolute coordinate string (e.g. B12 -> $B$12)
    /// Defaulting to true, absolute_column and absolute_row can optionally control
    /// whether the resulting cell_reference has an absolute column (e.g. B12 -> $B12)
    /// and absolute row (e.g. B12 -> B$12) respectively.
    /// </summary>
    /// <remarks>
    /// This is functionally equivalent to:
    /// cell_reference copy(*this);
    /// copy.column_absolute(absolute_column);
    /// copy.row_absolute(absolute_row);
    /// return copy;
    /// </remarks>
    cell_reference &make_absolute(bool absolute_column = true, bool absolute_row = true);

    /// <summary>
    /// Returns true if the reference refers to an absolute column, otherwise false.
    /// </summary>
    bool column_absolute() const;

    /// <summary>
    /// Makes this reference have an absolute column if absolute_column is true,
    /// otherwise not absolute.
    /// </summary>
    void column_absolute(bool absolute_column);

    /// <summary>
    /// Returns true if the reference refers to an absolute row, otherwise false.
    /// </summary>
    bool row_absolute() const;

    /// <summary>
    /// Makes this reference have an absolute row if absolute_row is true,
    /// otherwise not absolute.
    /// </summary>
    void row_absolute(bool absolute_row);

    // getters/setters

    /// <summary>
    /// Returns a string that identifies the column of this reference
    /// (e.g. second column from left is "B")
    /// </summary>
    column_t column() const;

    /// <summary>
    /// Sets the column of this reference from a string that identifies a particular column.
    /// </summary>
    void column(const std::string &column_string);

    /// <summary>
    /// Returns a 1-indexed numeric index of the column of this reference.
    /// </summary>
    column_t::index_t column_index() const;

    /// <summary>
    /// Sets the column of this reference from a 1-indexed number that identifies a particular column.
    /// </summary>
    void column_index(column_t column);

    /// <summary>
    /// Returns a 1-indexed numeric index of the row of this reference.
    /// </summary>
    row_t row() const;

    /// <summary>
    /// Sets the row of this reference from a 1-indexed number that identifies a particular row.
    /// </summary>
    void row(row_t row);

    /// <summary>
    /// Returns a cell_reference offset from this cell_reference by
    /// the number of columns and rows specified by the parameters.
    /// A negative value for column_offset or row_offset results
    /// in a reference above or left of this cell_reference, respectively.
    /// </summary>
    cell_reference make_offset(int column_offset, int row_offset) const;

    /// <summary>
    /// Returns a string like "A1" for cell_reference(1, 1).
    /// </summary>
    std::string to_string() const;

    /// <summary>
    /// Returns a 1x1 range_reference containing only this cell_reference.
    /// </summary>
    range_reference to_range() const;

    // operators

    /// <summary>
    /// I've always wanted to overload the comma operator.
    /// cell_reference("A", 1), cell_reference("B", 1) will return
    /// range_reference(cell_reference("A", 1), cell_reference("B", 1))
    /// </summary>
    range_reference operator,(const cell_reference &other) const;

    /// <summary>
    /// Returns true if this reference is identical to comparand including
    /// in absoluteness of column and row.
    /// </summary>
    bool operator==(const cell_reference &comparand) const;

    /// <summary>
    /// Constructs a cell_reference from reference_string and return the result
    /// of their comparison.
    /// </summary>
    bool operator==(const std::string &reference_string) const;

    /// <summary>
    /// Constructs a cell_reference from reference_string and return the result
    /// of their comparison.
    /// </summary>
    bool operator==(const char *reference_string) const;

    /// <summary>
    /// Returns true if this reference is not identical to comparand including
    /// in absoluteness of column and row.
    /// </summary>
    bool operator!=(const cell_reference &comparand) const;

    /// <summary>
    /// Constructs a cell_reference from reference_string and return the result
    /// of their comparison.
    /// </summary>
    bool operator!=(const std::string &reference_string) const;

    /// <summary>
    /// Constructs a cell_reference from reference_string and return the result
    /// of their comparison.
    /// </summary>
    bool operator!=(const char *reference_string) const;

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
    /// True if the reference's row is absolute. This looks like "A$1" in Excel.
    /// </summary>
    bool absolute_row_;

    /// <summary>
    /// True if the reference's column is absolute. This looks like "$A1" in Excel.
    /// </summary>
    bool absolute_column_;
};

} // namespace xlnt

namespace std {
template <>
struct hash<xlnt::cell_reference>
{
    size_t operator()(const xlnt::cell_reference &x) const
    {
        static_assert(std::is_same<decltype(x.row()), std::uint32_t>::value, "this hash function expects both row and column to be 32-bit numbers");
        static_assert(std::is_same<decltype(x.column_index()), std::uint32_t>::value, "this hash function expects both row and column to be 32-bit numbers");
        return hash<std::uint64_t>{}(x.row() | static_cast<std::uint64_t>(x.column_index()) << 32);
    }
};
} // namespace std
