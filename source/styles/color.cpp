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
#include <stdexcept>

#include <xlnt/styles/color.hpp>

namespace xlnt {

color::color()
{
}

color::color(type t, std::size_t v) : type_(t), index_(v)
{
}

color::color(type t, const std::string &v) : type_(t), rgb_string_(v)
{
}

void color::set_auto(std::size_t auto_index)
{
    type_ = type::auto_;
    index_ = auto_index;
}

void color::set_index(std::size_t index)
{
    type_ = type::indexed;
    index_ = index;
}

void color::set_theme(std::size_t theme)
{
    type_ = type::theme;
    index_ = theme;
}

color::type color::get_type() const
{
    return type_;
}

std::size_t color::get_auto() const
{
    if (type_ != type::auto_)
    {
        throw std::runtime_error("not auto color");
    }

    return index_;
}

std::size_t color::get_index() const
{
    if (type_ != type::indexed)
    {
        throw std::runtime_error("not indexed color");
    }

    return index_;
}

std::size_t color::get_theme() const
{
    if (type_ != type::theme)
    {
        throw std::runtime_error("not theme color");
    }

    return index_;
}

std::string color::get_rgb_string() const
{
    if (type_ != type::rgb)
    {
        throw std::runtime_error("not rgb color");
    }

    return rgb_string_;
}

std::string color::to_hash_string() const
{
    std::string hash_string = "color";
    
    hash_string.append(std::to_string(static_cast<std::size_t>(type_)));

    if (type_ != type::rgb)
    {
        hash_string.append(std::to_string(index_));
    }
    else
    {
        hash_string.append(rgb_string_);
    }

    return hash_string;
}

} // namespace xlnt
