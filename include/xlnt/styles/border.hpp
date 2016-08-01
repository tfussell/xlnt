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

#include <cstddef>
#include <functional>
#include <unordered_map>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/styles/border_style.hpp>
#include <xlnt/styles/color.hpp>
#include <xlnt/styles/diagonal_direction.hpp>
#include <xlnt/styles/side.hpp>
#include <xlnt/utils/hashable.hpp>

namespace xlnt {

enum class XLNT_CLASS border_side
{
	start,
	end,
	top,
	bottom,
	diagonal,
	vertical,
	horizontal
};

} // namespace xlnt

namespace std {

template<>
struct hash<xlnt::border_side>
{
	size_t operator()(const xlnt::border_side &k) const
	{
		return static_cast<std::size_t>(k);
	}
};

} // namepsace std

namespace xlnt {

/// <summary>
/// Describes the border style of a particular cell.
/// </summary>
class XLNT_CLASS border : public hashable
{
public:
	using side = border_side;

	class XLNT_CLASS border_property
	{
	public:
		bool has_color() const;
		const color &get_color() const;
		void set_color(const color &c);

		bool has_style() const;
		border_style get_style() const;
		void set_style(border_style style);

	private:
		bool has_color_ = false;
		color color_ = color::black();

		bool has_style_ = false;
		border_style style_;
	};

	static const std::unordered_map<side, std::string> &get_side_names();

	border();

	bool has_side(side s) const;
	const border_property &get_side(side s) const;
	void set_side(side s, const border_property &prop);

protected:
    std::string to_hash_string() const override;

private:
	std::unordered_map<side, border_property> sides_;
    bool outline_ = true;
    diagonal_direction diagonal_direction_ = diagonal_direction::neither;
};

} // namespace xlnt
