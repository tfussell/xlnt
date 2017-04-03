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
      size_(12.0)
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

font &font::outline(bool outline)
{
    outline_ = outline;
    return *this;
}

bool font::outline() const
{
    return outline_;
}

font &font::shadow(bool shadow)
{
    shadow_ = shadow;
    return *this;
}

bool font::shadow() const
{
    return shadow_;
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

bool font::has_charset() const
{
    return charset_.is_set();
}

font &font::charset(std::size_t charset)
{
    charset_ = charset;
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

bool font::operator==(const font &other) const
{
    if (bold() != other.bold())
    {
        return false;
    }

    if (has_color() != other.has_color())
    {
        return false;
    }

    if (has_color())
    {
        if (color() != other.color())
        {
            return false;
        }
    }

    if (has_family() != other.has_family())
    {
        return false;
    }

    if (has_family())
    {
        if (family() != other.family())
        {
            return false;
        }
    }

    if (italic() != other.italic())
    {
        return false;
    }

    if (name() != other.name())
    {
        return false;
    }

    if (has_scheme() != other.has_scheme())
    {
        return false;
    }

    if (has_scheme())
    {
        if (scheme() != other.scheme())
        {
            return false;
        }
    }

    if (std::fabs(size() - other.size()) != 0.0)
    {
        return false;
    }

    if (strikethrough() != other.strikethrough())
    {
        return false;
    }

    if (superscript() != other.superscript())
    {
        return false;
    }

    if (underline() != other.underline())
    {
        return false;
    }

    return true;
}

} // namespace xlnt
