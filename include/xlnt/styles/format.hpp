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
#include <xlnt/styles/base_format.hpp>

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
class XLNT_API format : public base_format
{
public:
	std::size_t id() const;

	bool has_style() const;
	void style(const std::string &name);
	void style(const xlnt::style &new_style);
	const class style &style() const;

	class border &border() override;
	const class border &border() const override;
	void border(const xlnt::border &new_border, bool applied) override;

	class fill &fill() override;
	const class fill &fill() const override;
	void fill(const xlnt::fill &new_fill, bool applied) override;
    
	class font &font() override;
	const class font &font() const override;
	void font(const xlnt::font &new_font, bool applied) override;

	class number_format &number_format() override;
	const class number_format &number_format() const override;
	void number_format(const xlnt::number_format &new_number_format, bool applied) override;

	format &alignment_id(std::size_t id);
	format &border_id(std::size_t id);
	format &fill_id(std::size_t id);
	format &font_id(std::size_t id);
	format &number_format_id(std::size_t id);
	format &protection_id(std::size_t id);

private:
	friend struct detail::stylesheet;
	format(detail::format_impl *d);
	detail::format_impl *d_;
};

} // namespace xlnt
