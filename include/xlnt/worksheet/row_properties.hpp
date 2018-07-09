// Copyright (c) 2014-2018 Thomas Fussell
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

#include <xlnt/xlnt_config.hpp>

namespace xlnt {

/// <summary>
/// The properties of a row in a worksheet.
/// </summary>
class XLNT_API row_properties
{
public:
    /// <summary>
    /// Row height
    /// </summary>
    optional<double> height;

    /// <summary>
    /// Distance in pixels from the bottom of the cell to the baseline of the cell content
    /// </summary>
    optional<double> dy_descent;

    /// <summary>
    /// Whether or not the height is different from the default
    /// </summary>
    bool custom_height = false;

    /// <summary>
    /// Whether or not the row should be hidden
    /// </summary>
    bool hidden = false;

    /// <summary>
    /// True if row style should be applied
    /// </summary>
    optional<bool> custom_format;

    /// <summary>
    /// The index to the style used by all cells in this row
    /// </summary>
    optional<std::size_t> style;
};

inline bool operator==(const row_properties &lhs, const row_properties &rhs)
{
    return lhs.height == rhs.height
        && lhs.dy_descent == rhs.dy_descent
        && lhs.custom_height == rhs.custom_height
        && lhs.hidden == rhs.hidden
        && lhs.custom_format == rhs.custom_format
        && lhs.style == rhs.style;
}

} // namespace xlnt
