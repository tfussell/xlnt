// Copyright (c) 2014-2016 Thomas Fussell
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

namespace detail {
struct format_impl;
struct stylesheet;
} // namespace detail

/// <summary>
/// Describes the formatting of a particular cell.
/// </summary>
class XLNT_API format
{
public:
	std::size_t id() const;

	class alignment &alignment();
	const class alignment &alignment() const;
	format alignment(const xlnt::alignment &new_alignment, bool applied);
    bool alignment_applied() const;

	class border &border();
	const class border &border() const;
	format border(const xlnt::border &new_border, bool applied);
    bool border_applied() const;

	class fill &fill();
	const class fill &fill() const;
	format fill(const xlnt::fill &new_fill, bool applied);
    bool fill_applied() const;
    
	class font &font();
	const class font &font() const;
	format font(const xlnt::font &new_font, bool applied);
    bool font_applied() const;

	class number_format &number_format();
	const class number_format &number_format() const;
	format number_format(const xlnt::number_format &new_number_format, bool applied);
    bool number_format_applied() const;

	class protection &protection();
	const class protection &protection() const;
	format protection(const xlnt::protection &new_protection, bool applied);
    bool protection_applied() const;

	bool has_style() const;
    void clear_style();
	format style(const std::string &name);
	format style(const xlnt::style &new_style);
	const class style style() const;

private:
	friend struct detail::stylesheet;
    friend class cell;
	format(detail::format_impl *d);
	detail::format_impl *d_;
};

} // namespace xlnt
