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
#include <xlnt/worksheet/page_margins.hpp>

namespace xlnt {

page_margins::page_margins() : default_(true), top_(0), left_(0), bottom_(0), right_(0), header_(0), footer_(0)
{
}

bool page_margins::is_default() const
{
    return default_;
}

double page_margins::get_top() const
{
    return top_;
}

void page_margins::set_top(double top)
{
    default_ = false;
    top_ = top;
}

double page_margins::get_left() const
{
    return left_;
}

void page_margins::set_left(double left)
{
    default_ = false;
    left_ = left;
}

double page_margins::get_bottom() const
{
    return bottom_;
}

void page_margins::set_bottom(double bottom)
{
    default_ = false;
    bottom_ = bottom;
}

double page_margins::get_right() const
{
    return right_;
}

void page_margins::set_right(double right)
{
    default_ = false;
    right_ = right;
}

double page_margins::get_header() const
{
    return header_;
}

void page_margins::set_header(double header)
{
    default_ = false;
    header_ = header;
}

double page_margins::get_footer() const
{
    return footer_;
}

void page_margins::set_footer(double footer)
{
    default_ = false;
    footer_ = footer;
}
    
} // namespace xlnt
