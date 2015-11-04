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

#include <string>

#include <xlnt/utils/hash_combine.hpp>

#include "xlnt_config.hpp"

namespace xlnt {

class XLNT_CLASS color
{
  public:
    enum class type
    {
        indexed,
        theme,
        auto_,
        rgb
    };

    static const color black();
    static const color white();
    static const color red();
    static const color darkred();
    static const color blue();
    static const color darkblue();
    static const color green();
    static const color darkgreen();
    static const color yellow();
    static const color darkyellow();

    color()
    {
    }

    color(type t, std::size_t v) : type_(t), index_(v)
    {
    }

    color(type t, const std::string &v) : type_(t), rgb_string_(v)
    {
    }

    void set_auto(std::size_t auto_index)
    {
        type_ = type::auto_;
        index_ = auto_index;
    }

    void set_index(std::size_t index)
    {
        type_ = type::indexed;
        index_ = index;
    }

    void set_theme(std::size_t theme)
    {
        type_ = type::theme;
        index_ = theme;
    }

    type get_type() const
    {
        return type_;
    }

    std::size_t get_auto() const
    {
        if (type_ != type::auto_)
        {
            throw std::runtime_error("not auto color");
        }

        return index_;
    }

    std::size_t get_index() const
    {
        if (type_ != type::indexed)
        {
            throw std::runtime_error("not indexed color");
        }

        return index_;
    }

    std::size_t get_theme() const
    {
        if (type_ != type::theme)
        {
            throw std::runtime_error("not theme color");
        }

        return index_;
    }

    std::string get_rgb_string() const
    {
        if (type_ != type::rgb)
        {
            throw std::runtime_error("not rgb color");
        }

        return rgb_string_;
    }

    std::size_t hash() const
    {
        auto seed = static_cast<std::size_t>(type_);

        if (type_ != type::rgb)
        {
            hash_combine(seed, index_);
        }
        else
        {
            hash_combine(seed, rgb_string_);
        }

        return seed;
    }

    bool operator==(const color &other) const
    {
        return hash() == other.hash();
    }

  private:
    type type_ = type::indexed;
    std::size_t index_ = 0;
    std::string rgb_string_;
};

} // namespace xlnt
