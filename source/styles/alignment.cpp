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

#include <xlnt/styles/alignment.hpp>

namespace xlnt {

optional<bool> alignment::wrap() const
{
    return wrap_text_;
}

alignment &alignment::wrap(bool wrap_text)
{
    wrap_text_ = wrap_text;
    return *this;
}

optional<bool> alignment::shrink() const
{
    return shrink_to_fit_;
}

alignment &alignment::shrink(bool shrink_to_fit)
{
    shrink_to_fit_ = shrink_to_fit;
    return *this;
}

optional<horizontal_alignment> alignment::horizontal() const
{
    return horizontal_;
}

alignment &alignment::horizontal(horizontal_alignment horizontal)
{
    horizontal_ = horizontal;
    return *this;
}

optional<vertical_alignment> alignment::vertical() const
{
    return vertical_;
}

alignment &alignment::vertical(vertical_alignment vertical)
{
    vertical_ = vertical;
    return *this;
}

alignment &alignment::indent(int value)
{
    indent_ = value;
    return *this;
}

optional<int> alignment::indent() const
{
    return indent_;
}

alignment &alignment::rotation(int value)
{
    text_rotation_ = value;
    return *this;
}

optional<int> alignment::rotation() const
{
    return text_rotation_;
}

XLNT_API bool operator==(const alignment &left, const alignment &right)
{
    if (left.horizontal().is_set() != right.horizontal().is_set())
    {
        return false;
    }

    if (left.horizontal().is_set())
    {
        if (left.horizontal().get() != right.horizontal().get())
        {
            return false;
        }
    }

    if (left.indent().is_set() != right.indent().is_set())
    {
        return false;
    }

    if (left.indent().is_set())
    {
        if (left.indent().get() != right.indent().get())
        {
            return false;
        }
    }

    if (left.rotation().is_set() != right.rotation().is_set())
    {
        return false;
    }

    if (left.rotation().is_set())
    {
        if (left.rotation().get() != right.rotation().get())
        {
            return false;
        }
    }

    if (left.shrink().is_set() != right.shrink().is_set())
    {
        return false;
    }

    if (left.shrink().is_set())
    {
        if (left.shrink().get() != right.shrink().get())
        {
            return false;
        }
    }

    if (left.vertical().is_set() != right.vertical().is_set())
    {
        return false;
    }

    if (left.vertical().is_set())
    {
        if (left.vertical().get() != right.vertical().get())
        {
            return false;
        }
    }

    if (left.wrap().is_set() != right.wrap().is_set())
    {
        return false;
    }

    if (left.wrap().is_set())
    {
        if (left.wrap().get() != right.wrap().get())
        {
            return false;
        }
    }

    return true;
}

} // namespace xlnt
