// Copyright (c) 2014-2018 Thomas Fussell
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
#include <xlnt/utils/optional.hpp>

namespace xlnt {

struct XLNT_API print_options
{
    /// <summary>
    /// if both grid_lines_set and this are true, grid lines are printed
    /// </summary>
    optional<bool> print_grid_lines;

    /// <summary>
    /// if both print grid lines and this are true, grid lines are printed
    /// </summary>
    optional<bool> grid_lines_set;

    /// <summary>
    /// print row and column headings
    /// </summary>
    optional<bool> print_headings;

    /// <summary>
    /// center on page horizontally
    /// </summary>
    optional<bool> horizontal_centered;

    /// <summary>
    /// center on page vertically
    /// </summary>
    optional<bool> vertical_centered;
};

inline bool operator==(const print_options& lhs, const print_options &rhs)
{
    return lhs.grid_lines_set == rhs.grid_lines_set
        && lhs.horizontal_centered == rhs.horizontal_centered
        && lhs.print_grid_lines == rhs.print_grid_lines
        && lhs.print_headings == rhs.print_headings
        && lhs.vertical_centered == rhs.vertical_centered;
}
} // namespace xlnt