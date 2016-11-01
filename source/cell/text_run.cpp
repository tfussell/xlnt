// Copyright (c) 2014-2016 Thomas Fussell
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

#include <xlnt/cell/text_run.hpp>
#include <xlnt/styles/color.hpp>

namespace xlnt {

text_run::text_run() : text_run("")
{
}

text_run::text_run(const std::string &string) : string_(string)
{
}

std::string text_run::get_string() const
{
	return string_;
}

void text_run::set_string(const std::string &string)
{
	string_ = string;
}

bool text_run::has_formatting() const
{
    return has_size() || has_color() || has_font() || has_family() || has_scheme();
}

bool text_run::has_size() const
{
    return size_.is_set();
}

std::size_t text_run::get_size() const
{
    return *size_;
}

void text_run::set_size(std::size_t size)
{
    size_ = size;
}

bool text_run::has_color() const
{
    return color_.is_set();
}

color text_run::get_color() const
{
    return *color_;
}

void text_run::set_color(const color &new_color)
{
    color_ = new_color;
}

bool text_run::has_font() const
{
    return font_.is_set();
}

std::string text_run::get_font() const
{
    return *font_;
}

void text_run::set_font(const std::string &font)
{
    font_ = font;
}

bool text_run::has_family() const
{
    return family_.is_set();
}

std::size_t text_run::get_family() const
{
    return *family_;
}

void text_run::set_family(std::size_t family)
{
    family_ = family;
}

bool text_run::has_scheme() const
{
    return scheme_.is_set();
}

std::string text_run::get_scheme() const
{
    return *scheme_;
}

void text_run::set_scheme(const std::string &scheme)
{
    scheme_ = scheme;
}

bool text_run::bold_set() const
{
    return bold_.is_set();
}

bool text_run::is_bold() const
{
    return *bold_;
}

void text_run::set_bold(bool bold)
{
    bold_ = bold;
}

} // namespace xlnt
