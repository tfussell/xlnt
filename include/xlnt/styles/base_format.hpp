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

#include <xlnt/xlnt_config.hpp>
#include <xlnt/styles/alignment.hpp>
#include <xlnt/styles/border.hpp>
#include <xlnt/styles/fill.hpp>
#include <xlnt/styles/font.hpp>
#include <xlnt/styles/number_format.hpp>
#include <xlnt/styles/protection.hpp>

namespace xlnt {

/// <summary>
/// Describes the formatting of a particular cell.
/// </summary>
class XLNT_CLASS base_format
{
public:
    // Alignment
    xlnt::alignment &alignment();
    const xlnt::alignment &alignment() const;
    void alignment(const xlnt::alignment &new_alignment, bool applied);
    bool alignment_applied() const;
    
    // Border
	xlnt::border &border();
    const xlnt::border &border() const;
    void border(const xlnt::border &new_border, bool applied);
    bool border_applied() const;
    
    // Fill
	xlnt::fill &fill();
    const xlnt::fill &fill() const;
    void fill(const xlnt::fill &new_fill, bool applied);
    bool fill_applied() const;
    
    // Font
	xlnt::font &font();
    const xlnt::font &font() const;
    void font(const xlnt::font &new_font, bool applied);
    bool font_applied() const;
    
    // Number Format
	virtual xlnt::number_format &number_format();
    virtual const xlnt::number_format &number_format() const;
    virtual void number_format(const xlnt::number_format &new_number_format, bool applied);
    bool number_format_applied() const;
    
    // Protection
    xlnt::protection &protection();
    const xlnt::protection &protection() const;
    void protection(const xlnt::protection &new_protection, bool applied);
    bool protection_applied() const;
    
protected:
	xlnt::alignment alignment_;
	xlnt::border border_;
	xlnt::fill fill_;
	xlnt::font font_;
	xlnt::number_format number_format_;
	xlnt::protection protection_;
    
    bool apply_alignment_ = false;
    bool apply_border_ = false;
    bool apply_fill_ = false;
    bool apply_font_ = false;
    bool apply_number_format_ = false;
    bool apply_protection_ = false;
};

} // namespace xlnt
