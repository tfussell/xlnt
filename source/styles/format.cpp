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

#include <detail/format_impl.hpp>
#include <detail/stylesheet.hpp>
#include <xlnt/styles/format.hpp>
#include <xlnt/styles/style.hpp>

namespace xlnt {

format::format(detail::format_impl *d) : d_(d)
{
}

std::size_t format::id() const
{
	return d_->id;
}

void format::style(const xlnt::style &new_style)
{
	d_->style = new_style.name();
}

void format::style(const std::string &new_style)
{
	for (auto &style : d_->parent->styles)
	{
		if (style.name() == new_style)
		{
			d_->style = new_style;
			return;
		}
	}

	throw key_not_found();
}

bool format::has_style() const
{
	return d_->style;
}

const style &format::style() const
{
	if (!has_style())
	{
		throw invalid_attribute();
	}

	return d_->parent->get_style(*d_->style);
}

xlnt::font &format::font()
{
	return base_format::font();
}

const xlnt::font &format::font() const
{
	return base_format::font();
}

void format::font(const xlnt::font &new_font, bool applied)
{
    font_id(d_->parent->add_font(new_font));
	base_format::font(new_font, applied);
}


xlnt::number_format &format::number_format()
{
	return base_format::number_format();
}

const xlnt::number_format &format::number_format() const
{
	return base_format::number_format();
}

void format::number_format(const xlnt::number_format &new_number_format, bool applied)
{
	auto copy = new_number_format;

	if (!copy.has_id())
	{
		copy.set_id(d_->parent->next_custom_number_format_id());
		d_->parent->number_formats.push_back(copy);
	}

	base_format::number_format(copy, applied);
}

format &format::font_id(std::size_t id)
{
    d_->font_id = id;
    return *this;
}

} // namespace xlnt
