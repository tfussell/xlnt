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
#include <xlnt/styles/common_style.hpp>

namespace xlnt {

common_style::common_style()
{
}

common_style::common_style(const common_style &other)
{
}

alignment &common_style::get_alignment()
{
    return alignment_;
}

const alignment &common_style::get_alignment() const
{
    return alignment_;
}

void common_style::set_alignment(const xlnt::alignment &new_alignment)
{
    alignment_ = new_alignment;
}

number_format &common_style::get_number_format()
{
    return number_format_;
}

const number_format &common_style::get_number_format() const
{
    return number_format_;
}

void common_style::set_number_format(const xlnt::number_format &new_number_format)
{
    number_format_ = new_number_format;
}

border &common_style::get_border()
{
    return border_;
}

const border &common_style::get_border() const
{
    return border_;
}

void common_style::set_border(const xlnt::border &new_border)
{
    border_ = new_border;
}

fill &common_style::get_fill()
{
    return fill_;
}

const fill &common_style::get_fill() const
{
    return fill_;
}

void common_style::set_fill(const xlnt::fill &new_fill)
{
    fill_ = new_fill;
}

font &common_style::get_font()
{
    return font_;
}

const font &common_style::get_font() const
{
    return font_;
}

void common_style::set_font(const xlnt::font &new_font)
{
    font_ = new_font;
}

protection &common_style::get_protection()
{
    return protection_;
}

const protection &common_style::get_protection() const
{
    return protection_;
}

void common_style::set_protection(const xlnt::protection &new_protection)
{
    protection_ = new_protection;
}

std::string common_style::to_hash_string() const
{
    return "common_style";
}

} // namespace xlnt
