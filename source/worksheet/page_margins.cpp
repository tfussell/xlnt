// Copyright (c) 2014-2017 Thomas Fussell
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

page_margins::page_margins()
{
}

double page_margins::top() const
{
    return top_;
}

void page_margins::top(double top)
{
    top_ = top;
}

double page_margins::left() const
{
    return left_;
}

void page_margins::left(double left)
{
    left_ = left;
}

double page_margins::bottom() const
{
    return bottom_;
}

void page_margins::bottom(double bottom)
{
    bottom_ = bottom;
}

double page_margins::right() const
{
    return right_;
}

void page_margins::right(double right)
{
    right_ = right;
}

double page_margins::header() const
{
    return header_;
}

void page_margins::header(double header)
{
    header_ = header;
}

double page_margins::footer() const
{
    return footer_;
}

void page_margins::footer(double footer)
{
    footer_ = footer;
}

} // namespace xlnt
