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

#include <xlnt/styles/font.hpp>

namespace xlnt {

void font::set_bold(bool bold)
{
    bold_ = bold;
}

bool font::is_bold() const
{
    return bold_;
}

void font::set_italic(bool italic)
{
    italic_ = italic;
}

bool font::is_italic() const
{
    return italic_;
}

void font::set_strikethrough(bool strikethrough)
{
    strikethrough_ = strikethrough;
}

bool font::is_strikethrough() const
{
    return strikethrough_;
}

void font::set_underline(underline_style new_underline)
{
    underline_ = new_underline;
}

bool font::is_underline() const
{
    return underline_ != underline_style::none;
}

font::underline_style font::get_underline() const
{
    return underline_;
}

void font::set_size(std::size_t size)
{
    size_ = size;
}

std::size_t font::get_size() const
{
    return size_;
}

void font::set_name(const std::string &name)
{
    name_ = name;
}
std::string font::get_name() const
{
    return name_;
}

void font::set_color(color c)
{
    color_ = c;
}

void font::set_family(std::size_t family)
{
    has_family_ = true;
    family_ = family;
}

void font::set_scheme(const std::string &scheme)
{
    has_scheme_ = true;
    scheme_ = scheme;
}

color font::get_color() const
{
    return color_;
}

bool font::has_family() const
{
    return has_family_;
}

std::size_t font::get_family() const
{
    return family_;
}

bool font::has_scheme() const
{
    return has_scheme_;
}

std::string font::to_hash_string() const
{
    std::string hash_string = "font";

    hash_string.append(std::to_string(bold_));
    hash_string.append(std::to_string(italic_));
    hash_string.append(std::to_string(superscript_));
    hash_string.append(std::to_string(subscript_));
    hash_string.append(std::to_string(strikethrough_));
    hash_string.append(name_);
    hash_string.append(std::to_string(size_));
    hash_string.append(std::to_string(static_cast<std::size_t>(underline_)));
    hash_string.append(std::to_string(color_.hash()));
    hash_string.append(std::to_string(family_));
    hash_string.append(scheme_);

    return hash_string;
}

} // namespace xlnt
