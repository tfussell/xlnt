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
#include <xlnt/styles/base_format.hpp>

namespace xlnt {

alignment &base_format::alignment()
{
    return alignment_;
}

const alignment &base_format::alignment() const
{
    return alignment_;
}

void base_format::alignment(const xlnt::alignment &new_alignment, bool apply)
{
    alignment_ = new_alignment;
	apply_alignment_ = apply;
}

number_format &base_format::number_format()
{
    return number_format_;
}

const number_format &base_format::number_format() const
{
    return number_format_;
}

void base_format::number_format(const xlnt::number_format &new_number_format, bool apply)
{
    number_format_ = new_number_format;
	apply_number_format_ = apply;
}

border &base_format::border()
{
    return border_;
}

const border &base_format::border() const
{
    return border_;
}

void base_format::border(const xlnt::border &new_border, bool apply)
{
    border_ = new_border;
	apply_border_ = apply;
}

fill &base_format::fill()
{
    return fill_;
}

const fill &base_format::fill() const
{
    return fill_;
}

void base_format::fill(const xlnt::fill &new_fill, bool apply)
{
    fill_ = new_fill;
	apply_fill_ = apply;
}

font &base_format::font()
{
    return font_;
}

const font &base_format::font() const
{
    return font_;
}

void base_format::font(const xlnt::font &new_font, bool apply)
{
    font_ = new_font;
	apply_font_ = apply;
}

protection &base_format::protection()
{
    return protection_;
}

const protection &base_format::protection() const
{
    return protection_;
}

void base_format::protection(const xlnt::protection &new_protection, bool apply)
{
	protection_ = new_protection;
	apply_protection_ = apply;
}

bool base_format::alignment_applied() const
{
    return apply_alignment_;
}

bool base_format::border_applied() const
{
    return apply_border_;
}

bool base_format::fill_applied() const
{
    return apply_fill_;
}

bool base_format::font_applied() const
{
    return apply_font_;
}

bool base_format::number_format_applied() const
{
    return apply_number_format_;
}

bool base_format::protection_applied() const
{
    return apply_protection_;
}

} // namespace xlnt
