// Copyright (c) 2014-2021 Thomas Fussell
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
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE
//
// @license: http://www.opensource.org/licenses/mit-license.php
// @author: see AUTHORS file

#pragma once

#include <xlnt/xlnt_config.hpp>

namespace xlnt {

/// <summary>
/// Properties applied to a column in a worksheet.
/// Columns can have a size and a style.
/// </summary>
class XLNT_API column_properties
{
public:
    /// <summary>
    /// The optional width of the column
    /// </summary>
    optional<double> width;

    /// <summary>
    /// If true, this is a custom width
    /// </summary>
    bool custom_width = false;

    /// <summary>
    /// The style index of this column. This shouldn't be used since style indices
    /// aren't supposed to be used directly in xlnt. (TODO)
    /// </summary>
    optional<std::size_t> style;

    /// <summary>
    /// Is this column sized to fit its content as best it can
    /// serialise if true
    /// </summary>
    bool best_fit = false;

    /// <summary>
    /// If true, this column will be hidden
    /// </summary>
    bool hidden = false;
};

inline bool operator==(const column_properties &lhs, const column_properties &rhs)
{
    return lhs.width == rhs.width
        && lhs.custom_width == rhs.custom_width
        && lhs.style == rhs.style
        && lhs.best_fit == rhs.best_fit
        && lhs.hidden == rhs.hidden;
}

} // namespace xlnt
