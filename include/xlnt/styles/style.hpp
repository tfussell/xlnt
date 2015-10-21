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

#include "alignment.hpp"
#include "border.hpp"
#include "fill.hpp"
#include "font.hpp"
#include "number_format.hpp"
#include "protection.hpp"

namespace xlnt {

class workbook;

class style
{
public:
    style();
    
    std::size_t hash() const;
    
    const alignment get_alignment() const;
    const border get_border() const;
    const fill get_fill() const;
    const font get_font() const;
    const number_format get_number_format() const;
    const protection get_protection() const;
    bool pivot_button() const;
    bool quote_prefix() const;
    
    std::size_t get_fill_index() const { return fill_index_; }
    std::size_t get_font_index() const { return font_index_; }
    std::size_t get_border_index() const { return border_index_; }
    std::size_t get_number_format_index() const { return number_format_index_; }
    
    bool operator==(const style &other) const
    {
        return hash() == other.hash();
    }
    
private:
    friend class workbook;
    
    std::size_t style_index_;
    
    std::size_t alignment_index_;
    alignment alignment_;
    
    std::size_t border_index_;
    border border_;
    
    std::size_t fill_index_;
    fill fill_;
    
    std::size_t font_index_;
    font font_;
    
    std::size_t number_format_index_;
    number_format number_format_;
    
    std::size_t protection_index_;
    protection protection_;
    
    bool pivot_button_;
    bool quote_prefix_;
};

} // namespace xlnt
