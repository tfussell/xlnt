// Copyright (c) 2018 Thomas Fussell
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
#include <xlnt/utils/numeric.hpp>
#include <xlnt/utils/optional.hpp>

namespace xlnt {

/// <summary>
/// General worksheet formatting properties.
/// </summary>
class XLNT_API sheet_format_properties
{
public:
    /// <summary>
    /// The base column width
    /// </summary>
    optional<double> base_col_width;

    /// <summary>
    /// The default row height is required
    /// </summary>
    double default_row_height = 15.0;

    /// <summary>
    /// The default column width
    /// </summary>
    optional<double> default_column_width;

    /// <summary>
    /// x14ac extension, dyDescent property
    /// </summary>
    optional<double> dy_descent;
};

inline bool operator==(const sheet_format_properties &lhs, const sheet_format_properties &rhs)
{
    return lhs.base_col_width == rhs.base_col_width
        && lhs.default_column_width == rhs.default_column_width
        && detail::float_equals(lhs.default_row_height, rhs.default_row_height)
        && lhs.dy_descent == rhs.dy_descent;
}

} // namespace xlnt
