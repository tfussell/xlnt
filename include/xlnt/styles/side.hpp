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

#include "../common/hash_combine.hpp"

namespace xlnt {

enum class border_style
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

class side
{
public:
    enum class color_type
    {
        theme,
        indexed
    };
    
    side();
    
    std::size_t hash() const
    {
        std::size_t seed = 0;
        
        if(style_assigned_)
        {
            hash_combine(seed, static_cast<std::size_t>(style_));
        }
        
        if(color_assigned_)
        {
            hash_combine(seed, static_cast<std::size_t>(color_type_));
            hash_combine(seed, color_);
        }
        
        return seed;
    }
    
    border_style get_style() const
    {
        return style_;
    }
    
    color_type get_color_type() const
    {
        return color_type_;
    }
    
    std::size_t get_color() const
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
        return other.style_ == style_ && color_type_ == other.color_type_ && color_ == other.color_;
    }
    
    void set_border_style(border_style bs)
    {
        style_assigned_ = true;
        style_ = bs;
    }
    
    void set_color(color_type type, std::size_t color)
    {
        color_assigned_ = true;
        color_type_ = type;
        color_ = color;
    }
    
private:
    bool style_assigned_ = false;
    border_style style_ = border_style::none;
    bool color_assigned_ = false;
    color_type color_type_ = color_type::indexed;
    std::size_t color_ = 0;
};
    
} // namespace xlnt
