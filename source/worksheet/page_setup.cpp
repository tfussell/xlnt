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
#include <xlnt/worksheet/page_setup.hpp>

namespace xlnt {

page_setup::page_setup()
    : break_(xlnt::page_break::none),
      sheet_state_(xlnt::sheet_state::visible),
      paper_size_(xlnt::paper_size::letter),
      orientation_(xlnt::orientation::portrait),
      fit_to_page_(false),
      fit_to_height_(false),
      fit_to_width_(false),
      horizontal_centered_(false),
      vertical_centered_(false),
      scale_(1)
{
}

page_break page_setup::page_break() const
{
    return break_;
}

void page_setup::page_break(xlnt::page_break b)
{
    break_ = b;
}

sheet_state page_setup::sheet_state() const
{
    return sheet_state_;
}

void page_setup::sheet_state(xlnt::sheet_state sheet_state)
{
    sheet_state_ = sheet_state;
}

paper_size page_setup::paper_size() const
{
    return paper_size_;
}

void page_setup::paper_size(xlnt::paper_size paper_size)
{
    paper_size_ = paper_size;
}

orientation page_setup::orientation() const
{
    return orientation_;
}

void page_setup::orientation(xlnt::orientation orientation)
{
    orientation_ = orientation;
}

bool page_setup::fit_to_page() const
{
    return fit_to_page_;
}

void page_setup::fit_to_page(bool fit_to_page)
{
    fit_to_page_ = fit_to_page;
}

bool page_setup::fit_to_height() const
{
    return fit_to_height_;
}

void page_setup::fit_to_height(bool fit_to_height)
{
    fit_to_height_ = fit_to_height;
}

bool page_setup::fit_to_width() const
{
    return fit_to_width_;
}

void page_setup::fit_to_width(bool fit_to_width)
{
    fit_to_width_ = fit_to_width;
}

void page_setup::horizontal_centered(bool horizontal_centered)
{
    horizontal_centered_ = horizontal_centered;
}

bool page_setup::horizontal_centered() const
{
    return horizontal_centered_;
}

void page_setup::vertical_centered(bool vertical_centered)
{
    vertical_centered_ = vertical_centered;
}

bool page_setup::vertical_centered() const
{
    return vertical_centered_;
}

void page_setup::scale(double scale)
{
    scale_ = scale;
}

double page_setup::scale() const
{
    return scale_;
}

} // namespace xlnt
