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

class workbook;

/// <summary>
/// Describes the entirety of the styling of a particular cell.
/// </summary>
class XLNT_CLASS style : public hashable
{
public:
    style();
    style(const style &other);
    style &operator=(const style &other);

    std::size_t hash() const;

    const alignment get_alignment() const;
    const border get_border() const;
    const fill get_fill() const;
    const font get_font() const;
    const number_format get_number_format() const;
    const protection get_protection() const;
    bool pivot_button() const;
    bool quote_prefix() const;

    std::size_t get_id() const;
    std::size_t get_fill_id() const;
    std::size_t get_font_id() const;
    std::size_t get_border_id() const;
    std::size_t get_number_format_id() const;

    void apply_alignment(bool apply);
    void apply_border(bool apply);
    void apply_fill(bool apply);
    void apply_font(bool apply);
    void apply_number_format(bool apply);
    void apply_protection(bool apply);
    
protected:
    std::string to_hash_string() const override;

private:
    friend class style_serializer;
    friend class workbook;

    std::size_t id_;

    bool alignment_apply_;
    alignment alignment_;

    bool border_apply_;
    std::size_t border_id_;
    border border_;

    bool fill_apply_;
    std::size_t fill_id_;
    fill fill_;

    bool font_apply_;
    std::size_t font_id_;
    font font_;

    bool number_format_apply_;
    std::size_t number_format_id_;
    number_format number_format_;

    bool protection_apply_;
    protection protection_;

    bool pivot_button_;
    bool quote_prefix_;
};

} // namespace xlnt
