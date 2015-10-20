// Copyright (c) 2015 Thomas Fussell
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

#include <xlnt/styles/color.hpp>

namespace xlnt {

template<typename T>
struct optional
{
    T value;
    bool initialized;
};
    
enum class border_style
{
    none,
    dashdot,
    dashdotdot,
    dashed,
    dotted,
    double_,
    hair,
    medium,
    mediumdashdot,
    mediumdashdotdot,
    mediumdashed,
    slantdashdot,
    thick,
    thin
};

enum class diagonal_direction
{
    none,
    up,
    down,
    both
};
    
class side
{
public:
    side(border_style style = border_style::none, color = color::black);
    
    bool operator==(const side &other) const
    {
        return other.style_ == style_;
    }
    
private:
    border_style style_;
    color color_;
};
    
class border
{
public:
    static border default_border();
    
    optional<side> start;
    optional<side> end;
    optional<side> left;
    optional<side> right;
    optional<side> top;
    optional<side> bottom;
    optional<side> diagonal;
    optional<side> vertical;
    optional<side> horizontal;

    bool outline;
    bool diagonal_up;
    bool diagonal_down;
    
    diagonal_direction diagonal_direction_;
    
    bool operator==(const border &other) const
    {
        return hash() == other.hash();
    }
    
    std::size_t hash() const { return 0; }
};

} // namespace xlnt
