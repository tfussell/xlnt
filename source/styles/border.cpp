// Copyright (c) 2014-2018 Thomas Fussell
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

#include <xlnt/styles/border.hpp>
#include <xlnt/utils/exceptions.hpp>
#include <detail/default_case.hpp>

namespace xlnt {

optional<xlnt::color> border::border_property::color() const
{
    return color_;
}

border::border_property &border::border_property::color(const xlnt::color &c)
{
    color_ = c;
    return *this;
}

optional<border_style> border::border_property::style() const
{
    return style_;
}

border::border_property &border::border_property::style(border_style s)
{
    style_ = s;
    return *this;
}

bool border::border_property::operator==(const border::border_property &right) const
{
    auto &left = *this;

    if (left.style().is_set() != right.style().is_set())
    {
        return false;
    }

    if (left.style().is_set())
    {
        if (left.style().get() != right.style().get())
        {
            return false;
        }
    }

    if (left.color().is_set() != right.color().is_set())
    {
        return false;
    }

    if (left.color().is_set())
    {
        if (left.color().get() != right.color().get())
        {
            return false;
        }
    }

    return true;
}

bool border::border_property::operator!=(const border::border_property &right) const
{
    return !(*this == right);
}

border::border()
{
}

const std::vector<xlnt::border_side> &border::all_sides()
{
    static auto *sides = new std::vector<xlnt::border_side>{xlnt::border_side::start, xlnt::border_side::end,
        xlnt::border_side::top, xlnt::border_side::bottom, xlnt::border_side::diagonal, xlnt::border_side::vertical,
        xlnt::border_side::horizontal};

    return *sides;
}

optional<border::border_property> border::side(border_side s) const
{
    switch (s)
    {
    case border_side::bottom:
        return bottom_;
    case border_side::top:
        return top_;
    case border_side::start:
        return start_;
    case border_side::end:
        return end_;
    case border_side::vertical:
        return vertical_;
    case border_side::horizontal:
        return horizontal_;
    case border_side::diagonal:
        return diagonal_;
    }

    default_case(start_);
}

border &border::side(border_side s, const border_property &prop)
{
    switch (s)
    {
    case border_side::bottom:
        bottom_ = prop;
        break;
    case border_side::top:
        top_ = prop;
        break;
    case border_side::start:
        start_ = prop;
        break;
    case border_side::end:
        end_ = prop;
        break;
    case border_side::vertical:
        vertical_ = prop;
        break;
    case border_side::horizontal:
        horizontal_ = prop;
        break;
    case border_side::diagonal:
        diagonal_ = prop;
        break;
    }

    return *this;
}

border &border::diagonal(diagonal_direction direction)
{
    diagonal_direction_ = direction;
    return *this;
}

optional<diagonal_direction> border::diagonal() const
{
    return diagonal_direction_;
}

bool border::operator==(const border &right) const
{
    auto &left = *this;

    for (auto side : border::all_sides())
    {
        if (left.side(side).is_set() != right.side(side).is_set())
        {
            return false;
        }

        if (left.side(side).is_set())
        {
            if (left.side(side).get() != right.side(side).get())
            {
                return false;
            }
        }
    }

    return true;
}

bool border::operator!=(const border &right) const
{
    return !(*this == right);
}

} // namespace xlnt
