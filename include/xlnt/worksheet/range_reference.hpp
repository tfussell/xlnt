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
    /// Convert a coordinate to an absolute coordinate string (B12 -> $B$12)
    /// </summary>
    static range_reference make_absolute(const range_reference &relative_reference);

    /// <summary>
    ///
    /// </summary>
    range_reference();

    /// <summary>
    ///
    /// </summary>
    explicit range_reference(const std::string &range_string);

    /// <summary>
    ///
    /// </summary>
    explicit range_reference(const char *range_string);

    /// <summary>
    ///
    /// </summary>
    explicit range_reference(const std::pair<cell_reference, cell_reference> &reference_pair);

    /// <summary>
    ///
    /// </summary>
    range_reference(const cell_reference &start, const cell_reference &end);

    /// <summary>
    ///
    /// </summary>
    range_reference(column_t column_index_start, row_t row_index_start, column_t column_index_end, row_t row_index_end);

    /// <summary>
    ///
    /// </summary>
    bool is_single_cell() const;

    /// <summary>
    ///
    /// </summary>
    std::size_t width() const;

    /// <summary>
    ///
    /// </summary>
    std::size_t height() const;

    /// <summary>
    ///
    /// </summary>
    cell_reference top_left() const;

    /// <summary>
    ///
    /// </summary>
    cell_reference bottom_right() const;

    /// <summary>
    ///
    /// </summary>
    cell_reference &top_left();

    /// <summary>
    ///
    /// </summary>
    cell_reference &bottom_right();

    /// <summary>
    ///
    /// </summary>
    range_reference make_offset(int column_offset, int row_offset) const;

    /// <summary>
    ///
    /// </summary>
    std::string to_string() const;

    /// <summary>
    ///
    /// </summary>
    bool operator==(const range_reference &comparand) const;

    /// <summary>
    ///
    /// </summary>
    bool operator==(const std::string &reference_string) const;

    /// <summary>
    ///
    /// </summary>
    bool operator==(const char *reference_string) const;

    /// <summary>
    ///
    /// </summary>
    bool operator!=(const range_reference &comparand) const;

    /// <summary>
    ///
    /// </summary>
    bool operator!=(const std::string &reference_string) const;

    /// <summary>
    ///
    /// </summary>
    bool operator!=(const char *reference_string) const;

    /// <summary>
    ///
    /// </summary>
    XLNT_API friend bool operator==(const std::string &reference_string, const range_reference &ref);

    /// <summary>
    ///
    /// </summary>
    XLNT_API friend bool operator==(const char *reference_string, const range_reference &ref);

    /// <summary>
    ///
    /// </summary>
    XLNT_API friend bool operator!=(const std::string &reference_string, const range_reference &ref);

    /// <summary>
    ///
    /// </summary>
    XLNT_API friend bool operator!=(const char *reference_string, const range_reference &ref);

private:
    /// <summary>
    ///
    /// </summary>
    cell_reference top_left_;

    /// <summary>
    ///
    /// </summary>
    cell_reference bottom_right_;
};

} // namespace xlnt
