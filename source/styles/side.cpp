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
#include <xlnt/styles/side.hpp>

namespace xlnt {

side::side()
{
}

std::string side::to_hash_string() const
{
    std::string hash_string;
    
    hash_string.append(border_style_ ? "1" : "0");
    
    if(border_style_)
    {
        hash_string.append(std::to_string(static_cast<std::size_t>(*border_style_)));
    }
    
    hash_string.append(color_ ? "1" : "0");
    
    if(color_)
    {
        hash_string.append(std::to_string(color_->hash()));
    }
    
    return hash_string;
}

std::experimental::optional<border_style> &side::get_border_style()
{
    return border_style_;
}

const std::experimental::optional<border_style> &side::get_border_style() const
{
    return border_style_;
}

std::experimental::optional<color> &side::get_color()
{
    return color_;
}

const std::experimental::optional<color> &side::get_color() const
{
    return color_;
}
    
} // namespace xlnt
