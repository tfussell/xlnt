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

namespace xlnt {
namespace detail {
    
struct cell_impl;
    
} // namespace detail    

class alignment;
class border;
class font;
class fill;
class number_format;
class protection;
    
class cell;

class style
{
public:
    const style default_style();
    
    style();
    style(const style &rhs);
    style &operator=(const style &rhs);
    
    style copy() const;
    
    std::size_t hash() const;
    
    const font &get_font() const;
    const fill &get_fill() const;
    const border &get_border() const;
    const alignment &get_alignment() const;
    const number_format &get_number_format() const;
    const protection &get_protection() const;
    bool pivot_button() const;
    bool quote_prefix() const;
    
private:
    cell get_parent();
    const cell get_parent() const;
    detail::cell_impl *parent_;
    std::size_t font_id_;
    std::size_t fill_id_;
    std::size_t border_id_;
    std::size_t alignment_id_;
    std::size_t number_format_id_;
    std::size_t protection_id_;
    std::size_t style_id_;
};

} // namespace xlnt
