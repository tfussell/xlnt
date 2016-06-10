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

class cell;

/// <summary>
/// Describes the formatting of a particular cell.
/// </summary>
class XLNT_CLASS base_format : public hashable
{
public:
    base_format();
    base_format(const base_format &other);
    base_format &operator=(const base_format &other);
    
    void reset();

    // Alignment
    alignment &get_alignment();
    const alignment &get_alignment() const;
    void set_alignment(const alignment &new_alignment);
    void remove_alignment();
    bool alignment_applied() const;
    void alignment_applied(bool applied);
    
    // Border
    border &get_border();
    const border &get_border() const;
    void set_border(const border &new_border);
    void remove_border();
    bool border_applied() const;
    void border_applied(bool applied);
    
    // Fill
    fill &get_fill();
    const fill &get_fill() const;
    void set_fill(const fill &new_fill);
    void remove_fill();
    bool fill_applied() const;
    void fill_applied(bool applied);
    
    // Font
    font &get_font();
    const font &get_font() const;
    void set_font(const font &new_font);
    void remove_font();
    bool font_applied() const;
    void font_applied(bool applied);
    
    // Number Format
    number_format &get_number_format();
    const number_format &get_number_format() const;
    void set_number_format(const number_format &new_number_format);
    void remove_number_format();
    bool number_format_applied() const;
    void number_format_applied(bool applied);
    
    // Protection
    protection &get_protection();
    const protection &get_protection() const;
    void set_protection(const protection &new_protection);
    void remove_protection();
    bool protection_applied() const;
    void protection_applied(bool applied);
    
protected:
    std::string to_hash_string() const override;

    alignment alignment_;
    border border_;
    fill fill_;
    font font_;
    number_format number_format_;
    protection protection_;
    
    bool apply_alignment_;
    bool apply_border_;
    bool apply_fill_;
    bool apply_font_;
    bool apply_number_format_;
    bool apply_protection_;
};

} // namespace xlnt
