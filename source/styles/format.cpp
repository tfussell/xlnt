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

void format::clear_style()
{
    d_->style.clear();
}

format format::style(const xlnt::style &new_style)
{
	d_ = d_->parent->find_or_create_with(d_, new_style.name());
    return format(d_);
}

format format::style(const std::string &new_style)
{
	d_->style = new_style;
    return format(d_);
}

bool format::has_style() const
{
	return d_->style;
}

const style format::style() const
{
	if (!has_style())
	{
		throw invalid_attribute();
	}

	return d_->parent->style(d_->style.get());
}

xlnt::alignment &format::alignment()
{
	return d_->parent->alignments.at(d_->alignment_id.get());
}

const xlnt::alignment &format::alignment() const
{
	return d_->parent->alignments.at(d_->alignment_id.get());
}

format format::alignment(const xlnt::alignment &new_alignment, bool applied)
{
    d_ = d_->parent->find_or_create_with(d_, new_alignment, applied);
    return format(d_);
}

xlnt::border &format::border()
{
	return d_->parent->borders.at(d_->border_id.get());
}

const xlnt::border &format::border() const
{
	return d_->parent->borders.at(d_->border_id.get());
}

format format::border(const xlnt::border &new_border, bool applied)
{
    d_ = d_->parent->find_or_create_with(d_, new_border, applied);
    return format(d_);
}

xlnt::fill &format::fill()
{
	return d_->parent->fills.at(d_->fill_id.get());
}

const xlnt::fill &format::fill() const
{
	return d_->parent->fills.at(d_->fill_id.get());
}

format format::fill(const xlnt::fill &new_fill, bool applied)
{
    d_ = d_->parent->find_or_create_with(d_, new_fill, applied);
    return format(d_);
}

xlnt::font &format::font()
{
	return d_->parent->fonts.at(d_->font_id.get());
}

const xlnt::font &format::font() const
{
	return d_->parent->fonts.at(d_->font_id.get());
}

format format::font(const xlnt::font &new_font, bool applied)
{
    d_ = d_->parent->find_or_create_with(d_, new_font, applied);
    return format(d_);
}

xlnt::number_format &format::number_format()
{
	return d_->parent->number_formats.at(d_->number_format_id.get());
}

const xlnt::number_format &format::number_format() const
{
	return d_->parent->number_formats.at(d_->number_format_id.get());
}

format format::number_format(const xlnt::number_format &new_number_format, bool applied)
{
	auto copy = new_number_format;

	if (!copy.has_id())
	{
		copy.id(d_->parent->next_custom_number_format_id());
		d_->parent->number_formats.push_back(copy);
	}

    d_ = d_->parent->find_or_create_with(d_, copy, applied);
    return format(d_);
}

xlnt::protection &format::protection()
{
	return d_->parent->protections.at(d_->protection_id.get());
}

const xlnt::protection &format::protection() const
{
	return d_->parent->protections.at(d_->protection_id.get());
}

format format::protection(const xlnt::protection &new_protection, bool applied)
{
    d_ = d_->parent->find_or_create_with(d_, new_protection, applied);
    return format(d_);
}

bool format::alignment_applied() const
{
    return d_->alignment_applied;
}

bool format::border_applied() const
{
    return d_->border_applied;
}

bool format::fill_applied() const
{
    return d_->fill_applied;
}

bool format::font_applied() const
{
    return d_->font_applied;
}

bool format::number_format_applied() const
{
    return d_->number_format_applied;
}

bool format::protection_applied() const
{
    return d_->protection_applied;
}

} // namespace xlnt
