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

#include "font.hpp"
#include "fill.hpp"
#include "borders.hpp"
#include "alignment.hpp"
#include "color.hpp"
#include "number_format.hpp"
#include "protection.hpp"

namespace xlnt {

class style
{
public:
    style(bool static_ = false) : static_(static_) {}
    style(const style &rhs);
    
    style copy() const;
    
    font get_font() const;
    void set_font(font font);
    
    fill get_fill() const;
    void set_fill(fill fill);
    
    borders get_borders() const;
    void set_borders(borders borders);
    
    alignment get_alignment() const;
    void set_alignment(alignment alignment);
    
    number_format &get_number_format() { return number_format_; }
    const number_format &get_number_format() const { return number_format_; }
    void set_number_format(number_format number_format);
    
    protection get_protection() const;
    void set_protection(protection protection);
    
private:
    bool static_ = false;
    font font_;
    fill fill_;
    borders borders_;
    alignment alignment_;
    number_format number_format_;
    protection protection_;
};

} // namespace xlnt
