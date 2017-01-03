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

#include <cmath>

#include <xlnt/styles/font.hpp>

namespace xlnt {

font::font()
    : name_("Calibri"),
      size_(12.0),
      bold_(false),
      italic_(false),
      superscript_(false),
      subscript_(false),
      strikethrough_(false),
      underline_(underline_style::none)
{
}

font &font::bold(bool bold)
{
    bold_ = bold;
    return *this;
}

bool font::bold() const
{
    return bold_;
}

font &font::superscript(bool superscript)
{
    superscript_ = superscript;
    return *this;
}

bool font::superscript() const
{
    return superscript_;
}

font &font::italic(bool italic)
{
    italic_ = italic;
    return *this;
}

bool font::italic() const
{
    return italic_;
}

font &font::strikethrough(bool strikethrough)
{
    strikethrough_ = strikethrough;
    return *this;
}

bool font::strikethrough() const
{
    return strikethrough_;
}

font &font::underline(underline_style new_underline)
{
    underline_ = new_underline;
    return *this;
}

bool font::underlined() const
{
    return underline_ != underline_style::none;
}

font::underline_style font::underline() const
{
    return underline_;
}

bool font::has_size() const
{
    return size_.is_set();
}

font &font::size(double size)
{
    size_ = size;
    return *this;
}

double font::size() const
{
    return size_.get();
}

bool font::has_name() const
{
    return name_.is_set();
}

font &font::name(const std::string &name)
{
    name_ = name;
    return *this;
}

std::string font::name() const
{
    return name_.get();
}

bool font::has_color() const
{
    return color_.is_set();
}

font &font::color(const xlnt::color &c)
{
    color_ = c;
    return *this;
}

bool font::has_family() const
{
    return family_.is_set();
}

font &font::family(std::size_t family)
{
    family_ = family;
    return *this;
}

bool font::has_scheme() const
{
    return scheme_.is_set();
}

font &font::scheme(const std::string &scheme)
{
    scheme_ = scheme;
    return *this;
}

color font::color() const
{
    return color_.get();
}

std::size_t font::family() const
{
    return family_.get();
}

std::string font::scheme() const
{
    return scheme_.get();
}

XLNT_API bool operator==(const font &left, const font &right)
{
    if (left.bold() != right.bold())
    {
        return false;
    }

    if (left.has_color() != right.has_color())
    {
        return false;
    }

    if (left.has_color())
    {
        if (left.color() != right.color())
        {
            return false;
        }
    }

    if (left.has_family() != right.has_family())
    {
        return false;
    }

    if (left.has_family())
    {
        if (left.family() != right.family())
        {
            return false;
        }
    }

    if (left.italic() != right.italic())
    {
        return false;
    }

    if (left.name() != right.name())
    {
        return false;
    }

    if (left.has_scheme() != right.has_scheme())
    {
        return false;
    }

    if (left.has_scheme())
    {
        if (left.scheme() != right.scheme())
        {
            return false;
        }
    }

    if (std::fabs(left.size() - right.size()) != 0.0)
    {
        return false;
    }

    if (left.strikethrough() != right.strikethrough())
    {
        return false;
    }

    if (left.superscript() != right.superscript())
    {
        return false;
    }

    if (left.underline() != right.underline())
    {
        return false;
    }

    return true;
}

} // namespace xlnt
