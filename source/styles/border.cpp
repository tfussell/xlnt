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

#include <xlnt/utils/exceptions.hpp>
#include <xlnt/styles/border.hpp>

namespace xlnt {

bool border::border_property::has_color() const
{
	return has_color_;
}

const color &border::border_property::get_color() const
{
	if (!has_color_)
	{
		throw invalid_attribute();
	}

	return color_;
}

void border::border_property::set_color(const color &c)
{
	has_color_ = true;
	color_ = c;
}

bool border::border_property::has_style() const
{
	return has_style_;
}

border_style border::border_property::get_style() const
{
	if (!has_style_)
	{
		throw invalid_attribute();
	}

	return style_;
}

void border::border_property::set_style(border_style s)
{
	has_style_ = true;
	style_ = s;
}

border border::default()
{
	return border();
}

border::border()
{
	sides_ =
	{
		{ xlnt::border::side::start, border_property() },
		{ xlnt::border::side::end, border_property() },
		{ xlnt::border::side::top, border_property() },
		{ xlnt::border::side::bottom, border_property() },
		{ xlnt::border::side::diagonal, border_property() }
	};
}

const std::unordered_map<xlnt::border::side, std::string> &border::get_side_names()
{
	static auto *sides =
		new std::unordered_map<xlnt::border::side, std::string>
		{
			{ xlnt::border::side::start, "left" },
			{ xlnt::border::side::end, "right" },
			{ xlnt::border::side::top, "top" },
			{ xlnt::border::side::bottom, "bottom" },
			{ xlnt::border::side::diagonal, "diagonal" },
			{ xlnt::border::side::vertical, "vertical" },
			{ xlnt::border::side::horizontal, "horizontal" }
		};

	return *sides;
}

bool border::has_side(side s) const
{
	return sides_.find(s) != sides_.end();
}

const border::border_property &border::get_side(side s) const
{
	if (has_side(s))
	{
		return sides_.at(s);
	}

	throw key_not_found();
}

void border::set_side(side s, const border_property &prop)
{
	sides_[s] = prop;
}

std::string border::to_hash_string() const
{
    std::string hash_string;

	for (const auto &side_name : get_side_names())
	{
		hash_string.append(side_name.second);

		if (has_side(side_name.first))
		{
			const auto &side_properties = get_side(side_name.first);

			if (side_properties.has_style())
			{
				hash_string.append(std::to_string(static_cast<std::size_t>(side_properties.get_style())));
			}
			else
			{
				hash_string.push_back('=');
			}

			if (side_properties.has_color())
			{
				hash_string.append(std::to_string(std::hash<xlnt::hashable>()(side_properties.get_color())));
			}
			else
			{
				hash_string.push_back('=');
			}
		}
		else
		{
			hash_string.push_back('=');
		}

		hash_string.push_back(' ');
	}

    return hash_string;
}

} // namespace xlnt
