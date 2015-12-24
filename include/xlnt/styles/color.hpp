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
#pragma once

#include <string>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/utils/hashable.hpp>

namespace xlnt {

/// <summary>
/// Colors can be applied to many parts of a cell's style.
/// </summary>
class XLNT_CLASS color : public hashable
{
public:
    /// <summary>
    /// Some colors are references to colors rather than having a particular RGB value.
    /// </summary>
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

    color();

    color(type t, std::size_t v);
    
    color(type t, const std::string &v);

    void set_auto(std::size_t auto_index);

    void set_index(std::size_t index);

    void set_theme(std::size_t theme);

    type get_type() const;

    std::size_t get_auto() const;

    std::size_t get_index() const;

    std::size_t get_theme() const;

    std::string get_rgb_string() const;
    
protected:
    std::string to_hash_string() const override;

private:
    type type_ = type::indexed;
    std::size_t index_ = 0;
    std::string rgb_string_;
};

} // namespace xlnt
