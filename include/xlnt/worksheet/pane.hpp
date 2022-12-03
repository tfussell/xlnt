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
#include <xlnt/cell/cell_reference.hpp>
#include <xlnt/cell/index_types.hpp>

namespace xlnt {

/// <summary>
/// Enumeration of possible states of a pane
/// </summary>
enum class XLNT_API pane_state
{
    frozen,
    frozen_split,
    split
};

/// <summary>
/// Enumeration of the four quadrants of a worksheet
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
/// Implements section 3.3.1.64 of ECMA 376 - Worksheet view pane
/// </summary>
struct XLNT_API pane
{
    /// <summary>
    /// The optional top left cell
    /// </summary>
    optional<cell_reference> top_left_cell;

    /// <summary>
    /// The state of the pane
    /// </summary>
    optional<pane_state> state = pane_state::split;

    /// <summary>
    /// The pane which contains the active cell
    /// </summary>
    optional<pane_corner> active_pane = pane_corner::top_left;

    /// <summary>
    /// The row where the split should take place
    /// </summary>
    optional<double> y_split = 0.0;

    /// <summary>
    /// The column where the split should take place
    /// </summary>
    optional<double> x_split = 0.0;

    /// <summary>
    /// Returns true if this pane is equal to rhs based on its top-left cell, state,
    /// active pane, and x/y split location.
    /// </summary>
    bool operator==(const pane &rhs) const
    {
        return top_left_cell == rhs.top_left_cell
            && state == rhs.state
            && active_pane == rhs.active_pane
            && y_split == rhs.y_split
            && x_split == rhs.x_split;
    }
};

} // namespace xlnt
