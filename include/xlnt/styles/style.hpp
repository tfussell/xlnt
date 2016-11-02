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

#include <cstdint>
#include <string>

#include <xlnt/xlnt_config.hpp>

namespace xlnt {

namespace detail {
struct style_impl;
struct stylesheet;
}

/// <summary>
/// Describes a style which has a name and can be applied to multiple individual
/// formats. In Excel this is a "Cell Style".
/// </summary>
class XLNT_API style
{
public:
    std::string name() const;
    style &name(const std::string &name);
    
    bool hidden() const;
	style &hidden(bool value);

	optional<bool> custom() const;
	style &custom(bool value);
    
    optional<std::size_t> builtin_id() const;
	style &builtin_id(std::size_t builtin_id);

    // Alignment

    /// <summary>
    /// Returns the alignment to be applied to contents of cells using this format.
    /// </summary>
    class alignment &alignment();

    /// <summary>
    /// Returns the alignment to be applied to contents of cells using this format.
    /// </summary>
    const class alignment &alignment() const;

    /// <summary>
    /// Sets the alignment of this format to new_alignment. If applied is true,
    /// the alignment will be used when displaying the cell.
    /// </summary>
    void alignment(const xlnt::alignment &new_alignment, bool applied = true, bool update = true);

    /// <summary>
    /// Returns true if the alignment is being applied.
    /// </summary>
    bool alignment_applied() const;
    
    // Border
    
    /// <summary>
    /// Returns the border properties to be applied to the outline of cells using this format.
    /// </summary>
    class border &border();

    /// <summary>
    /// Returns the border properties to be applied to the outline of cells using this format.
    /// </summary>
    const class border &border() const;

    /// <summary>
    /// Sets the border of this format to new_border. If applied is true,
    /// the border will be used when displaying the cell.
    /// </summary>
    void border(const xlnt::border &new_border, bool applied = true, bool update = true);

    /// <summary>
    /// Returns true if the border is being applied.
    /// </summary>
    bool border_applied() const;
    
    // Fill

    /// <summary>
    /// Returns the fill to be applied to the background of cells using this format.
    /// </summary>
    class fill &fill();

    /// <summary>
    /// Returns the fill to be applied to the background of cells using this format.
    /// </summary>
    const class fill &fill() const;

    /// <summary>
    /// Sets the fill of this format to new_fill and applies it if applied is true.
    /// </summary>
    void fill(const xlnt::fill &new_fill, bool applied = true, bool update = true);

    /// <summary>
    /// Returns true if the fill is being applied.
    /// </summary>
    bool fill_applied() const;
    
    // Font

    /// <summary>
    /// Returns the font to be applied to contents of cells using this format.
    /// </summary>
    class font &font();

    /// <summary>
    /// Returns the font to be applied to contents of cells using this format.
    /// </summary>
    const class font &font() const;

    /// <summary>
    /// Sets the font of this format to new_font and applies it if applied is true.
    /// </summary>
    void font(const xlnt::font &new_font, bool applied = true, bool update = true);

    /// <summary>
    /// Returns true if the font is being applied.
    /// </summary>
    bool font_applied() const;
    
    // Number Format

    /// <summary>
    /// Returns the number format to be applied to text/numbers in cells using this format.
    /// </summary>
    class number_format &number_format();

    /// <summary>
    /// Returns the number format to be applied to text/numbers in cells using this format.
    /// </summary>
    const class number_format &number_format() const;

    /// <summary>
    /// Sets the number format applied to text/numbers in cells using this format
    /// and applies it if applied is true.
    /// </summary>
    void number_format(const xlnt::number_format &new_number_format, bool applied = true, bool update = true);

    /// <summary>
    /// Returns true if the number format is being applied.
    /// </summary>
    bool number_format_applied() const;
    
    // Protection

    /// <summary>
    /// Returns the protection to be applied to the contents of cells using this format.
    /// </summary>
    class protection &protection();

    /// <summary>
    /// Returns the protection to be applied to the contents of cells using this format.
    /// </summary>
    const class protection &protection() const;

    /// <summary>
    /// Sets the protection of the contents of the cell using this format. Applies it if
    /// applied is true.
    /// </summary>
    void protection(const xlnt::protection &new_protection, bool applied = true, bool update = true);

    /// <summary>
    /// Returns true if the protection is being applied.
    /// </summary>
    bool protection_applied() const;

    /// <summary>
    /// Returns true if the formats are identical.
    /// </summary>
    XLNT_API friend bool operator==(const style &left, const style &right);

private:
	friend struct detail::stylesheet;

	style(detail::style_impl *d);

	detail::style_impl *d_;
};

} // namespace xlnt
