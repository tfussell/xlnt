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
#include <vector>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/styles/border_style.hpp>
#include <xlnt/styles/color.hpp>
#include <xlnt/styles/diagonal_direction.hpp>
#include <xlnt/styles/side.hpp>
#include <xlnt/utils/hashable.hpp>
#include <xlnt/utils/optional.hpp>

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

namespace xlnt {

/// <summary>
/// Describes the border style of a particular cell.
/// </summary>
class XLNT_CLASS border : public hashable
{
public:
	class XLNT_CLASS border_property
	{
	public:
		optional<class color> color() const;
		border_property &color(const xlnt::color &c);

		optional<border_style> style() const;
		border_property &style(border_style style);

	private:
		optional<class color> color_;
		optional<border_style> style_;
	};

	static const std::vector<border_side> &all_sides();

	border();

	optional<border_property> side(border_side s) const;
	border &side(border_side s, const border_property &prop);

	optional<diagonal_direction> diagonal() const;
	border &diagonal(diagonal_direction dir);

protected:
    std::string to_hash_string() const override;

private:
	optional<border_property> start_;
	optional<border_property> end_;
	optional<border_property> top_;
	optional<border_property> bottom_;
	optional<border_property> vertical_;
	optional<border_property> horizontal_;
	optional<border_property> diagonal_;

    //bool outline_ = true;
	optional<diagonal_direction> diagonal_direction_;
};

} // namespace xlnt
