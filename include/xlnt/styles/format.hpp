// Copyright (c) 2014-2017 Thomas Fussell
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

#include <cstddef>
#include <string>

#include <xlnt/xlnt_config.hpp>

namespace xlnt {

class alignment;
class border;
class cell;
class fill;
class font;
class number_format;
class protection;
class style;

namespace detail {

struct format_impl;
struct stylesheet;
class xlsx_producer;

} // namespace detail

/// <summary>
/// Describes the formatting of a particular cell.
/// </summary>
class XLNT_API format
{
public:
    /// <summary>
    /// Returns the alignment of this format.
    /// </summary>
    class alignment alignment() const;

    /// <summary>
    /// Sets the alignment of this format to new_alignment. Applied, which defaults
    /// to true, determines whether the alignment should be enabled for cells using
    /// this format.
    /// </summary>
    format alignment(const xlnt::alignment &new_alignment, bool applied = true);

    /// <summary>
    /// Returns true if the alignment of this format should be applied to cells
    /// using it.
    /// </summary>
    bool alignment_applied() const;

    /// <summary>
    /// Returns the border of this format.
    /// </summary>
    class border border() const;

    /// <summary>
    /// Sets the border of this format to new_border. Applied, which defaults
    /// to true, determines whether the border should be enabled for cells using
    /// this format.
    /// </summary>
    format border(const xlnt::border &new_border, bool applied = true);

    /// <summary>
    /// Returns true if the border set for this format should be applied to cells using the format.        
    /// </summary>
    bool border_applied() const;

    /// <summary>
    /// Returns the fill of this format.
    /// </summary>
    class fill fill() const;

    /// <summary>
    /// Sets the fill of this format to new_fill. Applied, which defaults
    /// to true, determines whether the border should be enabled for cells using
    /// this format.
    /// </summary>
    format fill(const xlnt::fill &new_fill, bool applied = true);

    /// <summary>
    /// Returns true if the fill set for this format should be applied to cells using the format.        
    /// </summary>
    bool fill_applied() const;

    /// <summary>
    /// Returns the font of this format.
    /// </summary>
    class font font() const;

    /// <summary>
    /// Sets the font of this format to new_font. Applied, which defaults
    /// to true, determines whether the font should be enabled for cells using
    /// this format.
    /// </summary>
    format font(const xlnt::font &new_font, bool applied = true);

    /// <summary>
    /// Returns true if the font set for this format should be applied to cells using the format.        
    /// </summary>
    bool font_applied() const;

    /// <summary>
    /// Returns the number format of this format.
    /// </summary>
    class number_format number_format() const;

    /// <summary>
    /// Sets the number format of this format to new_number_format. Applied, which defaults
    /// to true, determines whether the number format should be enabled for cells using
    /// this format.
    /// </summary>
    format number_format(const xlnt::number_format &new_number_format, bool applied = true);

    /// <summary>
    /// Returns true if the number_format set for this format should be applied to cells using the format.    
    /// </summary>
    bool number_format_applied() const;

    /// <summary>
    /// Returns the protection of this format.
    /// </summary>
    class protection protection() const;

    /// <summary>
    /// Returns true if the protection set for this format should be applied to cells using the format.
    /// </summary>
    bool protection_applied() const;

    /// <summary>
    /// Sets the protection of this format to new_protection. Applied, which defaults
    /// to true, determines whether the protection should be enabled for cells using
    /// this format.
    /// </summary>
    format protection(const xlnt::protection &new_protection, bool applied = true);

    /// <summary>
    /// Returns true if the pivot table interface is enabled for this format.
    /// </summary>
    bool pivot_button() const;

    /// <summary>
    /// If show is true, a pivot table interface will be displayed for cells using
	/// this format.
    /// </summary>
    void pivot_button(bool show);

    /// <summary>
    /// Returns true if this format should add a single-quote prefix for all text values.
    /// </summary>
    bool quote_prefix() const;

    /// <summary>
    /// If quote is true, enables a single-quote prefix for all text values in cells
	/// using this format (e.g. "abc" will appear as "'abc"). The text will also not
	/// be stored in sharedStrings when this is enabled.
    /// </summary>
    void quote_prefix(bool quote);

    /// <summary>
    /// Returns true if this format has a corresponding style applied.
    /// </summary>
    bool has_style() const;

    /// <summary>
    /// Removes the style from this format if it exists.
    /// </summary>
    void clear_style();

    /// <summary>
    /// Sets the style of this format to a style with the given name.
    /// </summary>
    format style(const std::string &name);

    /// <summary>
    /// Sets the style of this format to new_style.
    /// </summary>
    format style(const class style &new_style);

    /// <summary>
    /// Returns the style of this format. If it has no style, an invalid_parameter
    /// exception will be thrown.
    /// </summary>
    class style style();

    /// <summary>
    /// Returns the style of this format. If it has no style, an invalid_parameters
    /// exception will be thrown.
    /// </summary>
    const class style style() const;

private:
    friend struct detail::stylesheet;
    friend class detail::xlsx_producer;
    friend class cell;

    /// <summary>
    /// Constructs a format from an impl pointer.
    /// </summary>
    format(detail::format_impl *d);

    /// <summary>
    /// The internal implementation of this format
    /// </summary>
    detail::format_impl *d_;
};

} // namespace xlnt
