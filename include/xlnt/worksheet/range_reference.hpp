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
#include <xlnt/cell/cell_reference.hpp>

namespace xlnt {

/// <summary>
/// A range_reference describes a rectangular area of a worksheet with positive
/// width and height defined by a top-left and bottom-right corner.
/// </summary>
class XLNT_API range_reference
{
public:
    /// <summary>
    /// Converts relative reference coordinates to absolute coordinates (B12 -> $B$12)
    /// </summary>
    static range_reference make_absolute(const range_reference &relative_reference);

    /// <summary>
    /// Constructs a range reference equal to A1:A1
    /// </summary>
    range_reference();

    /// <summary>
    /// Constructs a range reference equivalent to the provided range_string in the form
    /// top_left:bottom_right.
    /// </summary>
    explicit range_reference(const std::string &range_string);

    /// <summary>
    /// Constructs a range reference equivalent to the provided range_string in the form
    /// top_left:bottom_right.
    /// </summary>
    explicit range_reference(const char *range_string);

    /// <summary>
    /// Constructs a range reference from cell references indicating top
    /// left and bottom right coordinates of the range.
    /// </summary>
    range_reference(const cell_reference &start, const cell_reference &end);

    /// <summary>
    /// Constructs a range reference from column and row indices.
    /// </summary>
    range_reference(column_t column_index_start, row_t row_index_start,
        column_t column_index_end, row_t row_index_end);

    /// <summary>
    /// Returns true if the range has a width and height of 1 cell.
    /// </summary>
    bool is_single_cell() const;

    /// <summary>
    /// Returns the number of columns encompassed by this range.
    /// </summary>
    std::size_t width() const;

    /// <summary>
    /// Returns the number of rows encompassed by this range.
    /// </summary>
    std::size_t height() const;

    /// <summary>
    /// Returns the coordinate of the top left cell of this range.
    /// </summary>
    cell_reference top_left() const;

    /// <summary>
    /// Returns the coordinate of the top right cell of this range.
    /// </summary>
    cell_reference top_right() const;

    /// <summary>
    /// Returns the coordinate of the bottom left cell of this range.
    /// </summary>
    cell_reference bottom_left() const;

    /// <summary>
    /// Returns the coordinate of the bottom right cell of this range.
    /// </summary>
    cell_reference bottom_right() const;

    /// <summary>
    /// Returns a new range reference with the same width and height as this
    /// range but shifted by the given number of columns and rows.
    /// </summary>
    range_reference make_offset(int column_offset, int row_offset) const;

    /// <summary>
    /// Returns a string representation of this range.
    /// </summary>
    std::string to_string() const;

    /// <summary>
    /// Returns true if this range is equivalent to the other range.
    /// </summary>
    bool operator==(const range_reference &comparand) const;

    /// <summary>
    /// Returns true if this range is equivalent to the string representation
    /// of the other range.
    /// </summary>
    bool operator==(const std::string &reference_string) const;

    /// <summary>
    /// Returns true if this range is equivalent to the string representation
    /// of the other range.
    /// </summary>
    bool operator==(const char *reference_string) const;

    /// <summary>
    /// Returns true if this range is not equivalent to the other range.
    /// </summary>
    bool operator!=(const range_reference &comparand) const;

    /// <summary>
    /// Returns true if this range is not equivalent to the string representation
    /// of the other range.
    /// </summary>
    bool operator!=(const std::string &reference_string) const;

    /// <summary>
    /// Returns true if this range is not equivalent to the string representation
    /// of the other range.
    /// </summary>
    bool operator!=(const char *reference_string) const;

private:
    /// <summary>
    /// The top left cell in the range
    /// </summary>
    cell_reference top_left_;

    /// <summary>
    /// The bottom right cell in the range
    /// </summary>
    cell_reference bottom_right_;
};

/// <summary>
/// Returns true if the string representation of the range is equivalent to ref.
/// </summary>
XLNT_API bool operator==(const std::string &reference_string, const range_reference &ref);

/// <summary>
/// Returns true if the string representation of the range is equivalent to ref.
/// </summary>
XLNT_API bool operator==(const char *reference_string, const range_reference &ref);

/// <summary>
/// Returns true if the string representation of the range is not equivalent to ref.
/// </summary>
XLNT_API bool operator!=(const std::string &reference_string, const range_reference &ref);

/// <summary>
/// Returns true if the string representation of the range is not equivalent to ref.
/// </summary>
XLNT_API bool operator!=(const char *reference_string, const range_reference &ref);

} // namespace xlnt
