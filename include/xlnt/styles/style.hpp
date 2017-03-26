// Copyright (c) 2014-2017 Thomas Fussell
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
#include <xlnt/utils/optional.hpp>

namespace xlnt {

class alignment;
class border;
class cell;
class fill;
class font;
class number_format;
class protection;

namespace detail {

struct style_impl;
struct stylesheet;
class xlsx_consumer;

} // namespace detail

/// <summary>
/// Describes a style which has a name and can be applied to multiple individual
/// formats. In Excel this is a "Cell Style".
/// </summary>
class XLNT_API style
{
public:
    /// <summary>
    /// Delete zero-argument constructor
    /// </summary>
    style() = delete;

    /// <summary>
    /// Default copy constructor. Constructs a style using the same PIMPL as other.
    /// </summary>
    style(const style &other) = default;

    /// <summary>
    /// Returns the name of this style.
    /// </summary>
    std::string name() const;

    /// <summary>
    /// Sets the name of this style to name.
    /// </summary>
    style name(const std::string &name);

    /// <summary>
    /// Returns true if this style is hidden.
    /// </summary>
    bool hidden() const;

    /// <summary>
    /// Sets the hidden state of this style to value. A hidden style will not 
	/// be shown in the list of selectable styles in the UI, but will still 
	/// apply its formatting to cells using it.
    /// </summary>
    style hidden(bool value);

    /// <summary>
    /// Returns true if this is a builtin style that has been customized and
	/// should therefore be persisted in the workbook.
    /// </summary>
    bool custom_builtin() const;

    /// <summary>
    /// Returns the index of the builtin style that this style is an instance 
	/// of or is a customized version thereof. If style::builtin() is false,
	/// this will throw an invalid_attribute exception.
    /// </summary>
    std::size_t builtin_id() const;

	/// <summary>
	/// Returns true if this is a builtin style.
	/// </summary>
	bool builtin() const;

    // Formatting (xf) components

    /// <summary>
    /// Returns the alignment of this style.
    /// </summary>
    class alignment alignment() const;

	/// <summary>
	/// Returns true if the alignment of this style should be applied to cells
	/// using it.
	/// </summary>
	bool alignment_applied() const;

    /// <summary>
    /// Sets the alignment of this style to new_alignment. Applied, which defaults
	/// to true, determines whether the alignment should be enabled for cells using
	/// this style.
    /// </summary>
    style alignment(const xlnt::alignment &new_alignment, bool applied = true);

    /// <summary>
    ///
    /// </summary>
    class border border() const;

	/// <summary>
	///
	/// </summary>
	bool border_applied() const;

    /// <summary>
    ///
    /// </summary>
    style border(const xlnt::border &new_border, bool applied = true);

    /// <summary>
    ///
    /// </summary>
    class fill fill() const;

	/// <summary>
	///
	/// </summary>
	bool fill_applied() const;

    /// <summary>
    ///
    /// </summary>
    style fill(const xlnt::fill &new_fill, bool applied = true);

    /// <summary>
    ///
    /// </summary>
    class font font() const;

	/// <summary>
	///
	/// </summary>
	bool font_applied() const;

    /// <summary>
    ///
    /// </summary>
    style font(const xlnt::font &new_font, bool applied = true);

    /// <summary>
    ///
    /// </summary>
    class number_format number_format() const;

	/// <summary>
	///
	/// </summary>
	bool number_format_applied() const;

    /// <summary>
    ///
    /// </summary>
    style number_format(const xlnt::number_format &new_number_format, bool applied = true);

    /// <summary>
    ///
    /// </summary>
    class protection protection() const;

	/// <summary>
	/// 
	/// </summary>
	bool protection_applied() const;

    /// <summary>
    /// 
    /// </summary>
    style protection(const xlnt::protection &new_protection, bool applied = true);

    /// <summary>
    /// Returns true if the pivot table interface is enabled for this style.
    /// </summary>
    bool pivot_button() const;

    /// <summary>
    /// If show is true, a pivot table interface will be displayed for cells using
	/// this style.
    /// </summary>
    void pivot_button(bool show);

    /// <summary>
    /// Returns true if this style should add a single-quote prefix for all text values.
    /// </summary>
    bool quote_prefix() const;

    /// <summary>
    /// If quote is true, enables a single-quote prefix for all text values in cells
	/// using this style (e.g. "abc" will appear as "'abc"). The text will also not
	/// be stored in sharedStrings when this is enabled.
    /// </summary>
    void quote_prefix(bool quote);

    /// <summary>
    /// Returns true if this style is equivalent to other.
    /// </summary>
    bool operator==(const style &other) const;

	/// <summary>
	/// Returns true if this style is not equivalent to other.
	/// </summary>
	bool operator!=(const style &other) const;

private:
    friend struct detail::stylesheet;
    friend class detail::xlsx_consumer;

    /// <summary>
    ///
    /// </summary>
    style(detail::style_impl *d);

    /// <summary>
    ///
    /// </summary>
    detail::style_impl *d_;
};

} // namespace xlnt
