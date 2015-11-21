// Copyright (c) 2014-2015 Thomas Fussell
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

#include <cstddef>
#include <functional>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/styles/diagonal_direction.hpp>
#include <xlnt/styles/side.hpp>
#include <xlnt/utils/hash_combine.hpp>

namespace xlnt {

/// <summary>
/// Describes the border style of a particular cell.
/// </summary>
class XLNT_CLASS border
{
  public:
    static border default_border();

    bool start_assigned = false;
    side start;
    bool end_assigned = false;
    side end;
    bool left_assigned = false;
    side left;
    bool right_assigned = false;
    side right;
    bool top_assigned = false;
    side top;
    bool bottom_assigned = false;
    side bottom;
    bool diagonal_assigned = false;
    side diagonal;
    bool vertical_assigned = false;
    side vertical;
    bool horizontal_assigned = false;
    side horizontal;

    bool outline = false;
    bool diagonal_up = false;
    bool diagonal_down = false;

    diagonal_direction diagonal_direction_ = diagonal_direction::none;

    bool operator==(const border &other) const;

    std::size_t hash() const;
};

} // namespace xlnt
