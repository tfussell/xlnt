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
class XLNT_CLASS format
{
public:
	std::size_t get_id() const;

	alignment get_alignment() const;
	void set_alignment(const alignment &new_alignment);

	border get_border() const;
	void set_border(const border &new_border);

	fill get_fill() const;
	void set_fill(const fill &new_fill);

	font get_font() const;
	void set_font(const font &new_font);

	number_format get_number_format() const;
	void set_number_format(const number_format &new_number_format);

	protection get_protection() const;
	void set_protection(const protection &new_protection);

	void set_style(const std::string &name);
	std::string get_name() const;

private:
	friend struct detail::stylesheet;
	format(detail::format_impl *d);
	detail::format_impl *d_;
};

} // namespace xlnt
