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
} // namespace detail

/// <summary>
/// Describes a style which has a name and can be applied to multiple individual
/// formats. In Excel this is a "Cell Style".
/// </summary>
class XLNT_API style
{
public:
    style() = delete;
    style(const style &other) = default;

    std::string name() const;
    style name(const std::string &name);

    class alignment &alignment();
	const class alignment &alignment() const;
	style alignment(const xlnt::alignment &new_alignment, bool applied = true);
    bool alignment_applied() const;

	class border &border();
	const class border &border() const;
	style border(const xlnt::border &new_border, bool applied = true);
    bool border_applied() const;

	class fill &fill();
	const class fill &fill() const;
	style fill(const xlnt::fill &new_fill, bool applied = true);
    bool fill_applied() const;
    
	class font &font();
	const class font &font() const;
	style font(const xlnt::font &new_font, bool applied = true);
    bool font_applied() const;

	class number_format &number_format();
	const class number_format &number_format() const;
	style number_format(const xlnt::number_format &new_number_format, bool applied = true);
    bool number_format_applied() const;

	class protection &protection();
	const class protection &protection() const;
	style protection(const xlnt::protection &new_protection, bool applied = true);
    bool protection_applied() const;
    
    bool hidden() const;
	style hidden(bool value);

	optional<bool> custom() const;
	style custom(bool value);
    
    optional<std::size_t> builtin_id() const;
	style builtin_id(std::size_t builtin_id);

	bool operator==(const style &other) const;

private:
	friend struct detail::stylesheet;
	style(detail::style_impl *d);
	detail::style_impl *d_;
};

} // namespace xlnt
