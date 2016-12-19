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

#include <xlnt/xlnt_config.hpp>
#include <xlnt/cell/cell_reference.hpp>
#include <xlnt/cell/index_types.hpp>

namespace xlnt {

/// <summary>
///
/// </summary>
enum class XLNT_API pane_state
{
    frozen,
    frozen_split,
    split
};

/// <summary>
///
/// </summary>
enum class XLNT_API pane_corner
{
    top_left,
    top_right,
    bottom_left,
    bottom_right
};

/// <summary>
/// A fixed portion of a worksheet.
/// </summary>
struct XLNT_API pane
{
    /// <summary>
    ///
    /// </summary>
    optional<cell_reference> top_left_cell;

    /// <summary>
    ///
    /// </summary>
    pane_state state = pane_state::split;

    /// <summary>
    ///
    /// </summary>
    pane_corner active_pane = pane_corner::top_left;

    /// <summary>
    ///
    /// </summary>
    row_t y_split = 1;

    /// <summary>
    ///
    /// </summary>
    column_t x_split = 1;

    bool operator==(const pane &rhs) const
    {
        return top_left_cell == rhs.top_left_cell && state == rhs.state && active_pane == rhs.active_pane
            && y_split == rhs.y_split && x_split == rhs.x_split;
    }
};

} // namespace xlnt
