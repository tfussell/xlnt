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
#include <xlnt/cell/cell.hpp>
#include <xlnt/styles/font.hpp>
#include <xlnt/styles/number_format.hpp>
#include <xlnt/styles/protection.hpp>
#include <xlnt/styles/style.hpp>

namespace {

template <class T>
void hash_combine(std::size_t &seed, const T &v)
{
    std::hash<T> hasher;
    seed ^= hasher(v) + 0x9e3779b9 + (seed << 6) + (seed >> 2);
}
}

namespace xlnt {

const color color::black()
{
    return color(color::type::rgb, "ff000000");
}

const color color::white()
{
    return color(color::type::rgb, "ffffffff");
}

style::style()
    : id_(0),
      alignment_apply_(false),
      border_apply_(false),
      border_id_(0),
      fill_apply_(false),
      fill_id_(0),
      font_apply_(false),
      font_id_(0),
      number_format_apply_(false),
      number_format_id_(0),
      protection_apply_(false),
      pivot_button_(false),
      quote_prefix_(false)
{
}

style::style(const style &other)
    : id_(other.id_),
      alignment_apply_(other.alignment_apply_),
      alignment_(other.alignment_),
      border_apply_(other.border_apply_),
      border_id_(other.border_id_),
      border_(other.border_),
      fill_apply_(other.fill_apply_),
      fill_id_(other.fill_id_),
      fill_(other.fill_),
      font_apply_(other.font_apply_),
      font_id_(other.font_id_),
      font_(other.font_),
      number_format_apply_(other.number_format_apply_),
      number_format_id_(other.number_format_id_),
      number_format_(other.number_format_),
      protection_apply_(other.protection_apply_),
      protection_(other.protection_),
      pivot_button_(other.pivot_button_),
      quote_prefix_(other.quote_prefix_)
{
}

style &style::operator=(const style &other)
{
    id_ = other.id_;
    alignment_ = other.alignment_;
    alignment_apply_ = other.alignment_apply_;
    border_ = other.border_;
    border_apply_ = other.border_apply_;
    border_id_ = other.border_id_;
    fill_ = other.fill_;
    fill_apply_ = other.fill_apply_;
    fill_id_ = other.fill_id_;
    font_ = other.font_;
    font_apply_ = other.font_apply_;
    font_id_ = other.font_id_;
    number_format_ = other.number_format_;
    number_format_apply_ = other.number_format_apply_;
    number_format_id_ = other.number_format_id_;
    protection_ = other.protection_;
    protection_apply_ = other.protection_apply_;
    pivot_button_ = other.pivot_button_;
    quote_prefix_ = other.quote_prefix_;

    return *this;
}

std::string style::to_hash_string() const
{
    std::string hash_string("style");
    
    hash_string.append(std::to_string(alignment_apply_));
    hash_string.append(alignment_apply_ ? std::to_string(alignment_.hash()) : " ");

    hash_string.append(std::to_string(border_apply_));
    hash_string.append(border_apply_ ? std::to_string(border_id_) : " ");

    hash_string.append(std::to_string(font_apply_));
    hash_string.append(font_apply_ ? std::to_string(font_id_) : " ");

    hash_string.append(std::to_string(fill_apply_));
    hash_string.append(fill_apply_ ? std::to_string(fill_id_) : " ");

    hash_string.append(std::to_string(number_format_apply_));
    hash_string.append(number_format_apply_ ? std::to_string(number_format_id_) : " ");

    hash_string.append(std::to_string(protection_apply_));
    hash_string.append(protection_apply_ ? std::to_string(protection_.hash()) : " ");

    return hash_string;
}

const number_format style::get_number_format() const
{
    return number_format_;
}

const border style::get_border() const
{
    return border_;
}

const alignment style::get_alignment() const
{
    return alignment_;
}

const fill style::get_fill() const
{
    return fill_;
}

const font style::get_font() const
{
    return font_;
}

const protection style::get_protection() const
{
    return protection_;
}

bool style::pivot_button() const
{
    return pivot_button_;
}

bool style::quote_prefix() const
{
    return quote_prefix_;
}

std::size_t style::get_id() const
{
    return id_;
}

std::size_t style::get_fill_id() const
{
    return fill_id_;
}

std::size_t style::get_font_id() const
{
    return font_id_;
}

std::size_t style::get_border_id() const
{
    return border_id_;
}

std::size_t style::get_number_format_id() const
{
    return number_format_id_;
}

void style::apply_alignment(bool apply)
{
    alignment_apply_ = apply;
}

void style::apply_border(bool apply)
{
    border_apply_ = apply;
}

void style::apply_fill(bool apply)
{
    fill_apply_ = apply;
}

void style::apply_font(bool apply)
{
    font_apply_ = apply;
}

void style::apply_number_format(bool apply)
{
    number_format_apply_ = apply;
}

void style::apply_protection(bool apply)
{
    protection_apply_ = apply;
}

} // namespace xlnt
