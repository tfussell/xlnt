// Copyright (c) 2015 Thomas Fussell
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
#pragma once

#include <cstddef>

#include <xlnt/styles/color.hpp>

#include <xlnt/xlnt_config.hpp>

namespace xlnt {

enum class XLNT_CLASS border_style
{
    none,
    dashdot,
    dashdotdot,
    dashed,
    dotted,
    double_,
    hair,
    medium,
    mediumdashdot,
    mediumdashdotdot,
    mediumdashed,
    slantdashdot,
    thick,
    thin
};

class XLNT_CLASS side
{
  public:
    side();

    std::size_t hash() const
    {
        std::size_t seed = 0;

        hash_combine(seed, style_assigned_);

        if (style_assigned_)
        {
            hash_combine(seed, static_cast<std::size_t>(style_));
        }

        hash_combine(seed, color_assigned_);

        if (color_assigned_)
        {
            hash_combine(seed, color_.hash());
        }

        return seed;
    }

    border_style get_style() const
    {
        return style_;
    }

    color get_color() const
    {
        return color_;
    }

    bool is_style_assigned() const
    {
        return style_assigned_;
    }

    bool is_color_assigned() const
    {
        return color_assigned_;
    }

    bool operator==(const side &other) const
    {
        return other.style_ == style_ && color_ == other.color_;
    }

    void set_border_style(border_style bs)
    {
        style_assigned_ = true;
        style_ = bs;
    }

    void set_color(color new_color)
    {
        color_assigned_ = true;
        color_ = new_color;
    }

  private:
    bool style_assigned_ = false;
    border_style style_ = border_style::none;
    bool color_assigned_ = false;
    color color_ = color::black();
};

} // namespace xlnt
