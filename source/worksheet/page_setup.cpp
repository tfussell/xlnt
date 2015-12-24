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
#include <xlnt/worksheet/page_setup.hpp>

namespace xlnt {

page_setup::page_setup()
    : default_(true),
      break_(page_break::none),
      sheet_state_(sheet_state::visible),
      paper_size_(paper_size::letter),
      orientation_(orientation::portrait),
      fit_to_page_(false),
      fit_to_height_(false),
      fit_to_width_(false),
      horizontal_centered_(false),
      vertical_centered_(false),
      scale_(1)
{
}

bool page_setup::is_default() const
{
    return default_;
}

page_break page_setup::get_break() const
{
    return break_;
}

void page_setup::set_break(page_break b)
{
    default_ = false;
    break_ = b;
}

sheet_state page_setup::get_sheet_state() const
{
    return sheet_state_;
}

void page_setup::set_sheet_state(sheet_state sheet_state)
{
    sheet_state_ = sheet_state;
}

paper_size page_setup::get_paper_size() const
{
    return paper_size_;
}

void page_setup::set_paper_size(paper_size paper_size)
{
    default_ = false;
    paper_size_ = paper_size;
}

orientation page_setup::get_orientation() const
{
    return orientation_;
}

void page_setup::set_orientation(orientation orientation)
{
    default_ = false;
    orientation_ = orientation;
}

bool page_setup::fit_to_page() const
{
    return fit_to_page_;
}

void page_setup::set_fit_to_page(bool fit_to_page)
{
    default_ = false;
    fit_to_page_ = fit_to_page;
}

bool page_setup::fit_to_height() const
{
    return fit_to_height_;
}

void page_setup::set_fit_to_height(bool fit_to_height)
{
    default_ = false;
    fit_to_height_ = fit_to_height;
}

bool page_setup::fit_to_width() const
{
    return fit_to_width_;
}

void page_setup::set_fit_to_width(bool fit_to_width)
{
    default_ = false;
    fit_to_width_ = fit_to_width;
}

void page_setup::set_horizontal_centered(bool horizontal_centered)
{
    default_ = false;
    horizontal_centered_ = horizontal_centered;
}

bool page_setup::get_horizontal_centered() const
{
    return horizontal_centered_;
}

void page_setup::set_vertical_centered(bool vertical_centered)
{
    default_ = false;
    vertical_centered_ = vertical_centered;
}

bool page_setup::get_vertical_centered() const
{
    return vertical_centered_;
}

void page_setup::set_scale(double scale)
{
    default_ = false;
    scale_ = scale;
}

double page_setup::get_scale() const
{
    return scale_;
}

} // namespace xlnt
