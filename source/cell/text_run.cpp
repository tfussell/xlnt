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

std::string text_run::string() const
{
	return string_;
}

void text_run::string(const std::string &string)
{
	string_ = string;
}

bool text_run::has_formatting() const
{
    return font_.is_set();
}

bool text_run::has_size() const
{
    return font_.is_set() && font_.get().has_size();
}

std::size_t text_run::size() const
{
    return font_.get().size();
}

void text_run::size(std::size_t size)
{
    if (!font_.is_set())
    {
        font_ = xlnt::font();
    }

    font_.get().size(size);
}

bool text_run::has_color() const
{
    return font_.is_set() && font_.get().has_color();
}

color text_run::color() const
{
    return font_.get().color();
}

void text_run::color(const class color &new_color)
{
    if (!font_.is_set())
    {
        font_ = xlnt::font();
    }

    font_.get().color(new_color);
}

bool text_run::has_font() const
{
    return font_.is_set() && font_.get().has_name();
}

std::string text_run::font() const
{
    return font_.get().name();
}

void text_run::font(const std::string &font)
{
    if (!font_.is_set())
    {
        font_ = xlnt::font();
    }

    font_.get().name(font);
}

bool text_run::has_family() const
{
    return font_.is_set() && font_.get().has_family();
}

std::size_t text_run::family() const
{
    return font_.get().family();
}

void text_run::family(std::size_t family)
{
    if (!font_.is_set())
    {
        font_ = xlnt::font();
    }

    font_.get().family(family);
}

bool text_run::has_scheme() const
{
    return font_.is_set() && font_.get().has_scheme();
}

std::string text_run::scheme() const
{
    return font_.get().scheme();
}

void text_run::scheme(const std::string &scheme)
{
    if (!font_.is_set())
    {
        font_ = xlnt::font();
    }

    font_.get().scheme(scheme);
}

bool text_run::bold() const
{
    return font_.is_set() && font_.get().bold();
}

void text_run::bold(bool bold)
{
    if (!font_.is_set())
    {
        font_ = xlnt::font();
    }

    font_.get().bold(bold);
}

void text_run::underline(font::underline_style style)
{
    if (!font_.is_set())
    {
        font_ = xlnt::font();
    }

    font_.get().underline(style);
}

} // namespace xlnt
