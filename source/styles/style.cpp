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

#include <detail/style_impl.hpp> // include order is important here
#include <detail/stylesheet.hpp>
#include <xlnt/styles/alignment.hpp>
#include <xlnt/styles/border.hpp>
#include <xlnt/styles/fill.hpp>
#include <xlnt/styles/font.hpp>
#include <xlnt/styles/number_format.hpp>
#include <xlnt/styles/protection.hpp>
#include <xlnt/styles/style.hpp>

namespace xlnt {

style::style(detail::style_impl *d) : d_(d)
{
}

bool style::hidden() const
{
    return d_->hidden_style;
}

style &style::hidden(bool value)
{
    d_->hidden_style = value;
	return *this;
}

optional<std::size_t> style::builtin_id() const
{
    return d_->builtin_id;
}

style &style::builtin_id(std::size_t builtin_id)
{
    d_->builtin_id = builtin_id;
	return *this;
}

std::string style::name() const
{
    return d_->name;
}

style &style::name(const std::string &name)
{
    d_->name = name;
	return *this;
}

optional<bool> style::custom() const
{
	return d_->custom_builtin;
}

style &style::custom(bool value)
{
	d_->custom_builtin = value;
	return *this;
}

alignment &style::alignment()
{
    return d_->parent->alignments.at(d_->alignment_id);
}

const alignment &style::alignment() const
{
    return d_->parent->alignments.at(d_->alignment_id);
}

void style::alignment(const xlnt::alignment &new_alignment, bool applied, bool update)
{
    if (update)
    {
        d_->alignment_id = 0;
    }
}

xlnt::border &style::border()
{
    const auto &self = *this;
    return const_cast<class border &>(self.border());
}

const xlnt::border &style::border() const
{
    if (!d_->border_id)
    {
        throw xlnt::exception("no border");
    }

	return d_->parent->borders.at(d_->border_id.get());
}

void style::border(const xlnt::border &new_border, bool applied, bool update)
{
    d_->border_id = d_->parent->add_border(new_border);
}

xlnt::fill &style::fill()
{
    const auto &self = *this;
    return const_cast<class fill &>(self.fill());
}

const xlnt::fill &style::fill() const
{
    if (!d_->fill_id)
    {
        throw xlnt::exception("no border");
    }

	return d_->parent->fills.at(d_->fill_id.get());
}

void style::fill(const xlnt::fill &new_fill, bool applied, bool update)
{
    d_->fill_id = d_->parent->add_fill(new_fill);
    d_->fill_applied = applied;
}

xlnt::font &style::font()
{
    const auto &self = *this;
    return const_cast<class font &>(self.font());
}

const xlnt::font &style::font() const
{
    if (!d_->font_id)
    {
        throw xlnt::exception("no border");
    }

	return d_->parent->fonts.at(d_->font_id.get());
}

void style::font(const xlnt::font &new_font, bool applied, bool update)
{
    d_->font_id = d_->parent->add_font(new_font);
    d_->font_applied = applied;
}

xlnt::number_format &style::number_format()
{
    const auto &self = *this;
    return const_cast<class number_format &>(self.number_format());
}

const xlnt::number_format &style::number_format() const
{
    if (!d_->number_format_id)
    {
        throw xlnt::exception("no number format");
    }

	return d_->parent->number_formats.at(d_->number_format_id.get());
}

void style::number_format(const xlnt::number_format &new_number_format, bool applied, bool update)
{
	auto copy = new_number_format;

	if (!copy.has_id())
	{
		copy.set_id(d_->parent->next_custom_number_format_id());
		d_->parent->number_formats[copy.get_id()] = copy;
	}

    d_->number_format_id = copy.get_id();
    d_->number_format_applied = applied;
}

protection &style::protection()
{
    return d_->parent->protections.at(d_->protection_id);
}

const protection &style::protection() const
{
    return d_->parent->protections.at(d_->protection_id);
}

void style::protection(const xlnt::protection &new_protection, bool applied, bool update)
{
    if (update)
    {
        d_->protection_id = 0;
    }
}


bool style::alignment_applied() const
{
    return d_->alignment_applied;
}

bool style::border_applied() const
{
    return d_->border_applied;
}

bool style::fill_applied() const
{
    return d_->fill_applied;
}

bool style::font_applied() const
{
    return d_->font_applied;
}

bool style::number_format_applied() const
{
    return d_->number_format_applied;
}

bool style::protection_applied() const
{
    return d_->protection_applied;
}

XLNT_API bool operator==(const style &left, const style &right)
{
	return left.d_ == right.d_;
}

} // namespace xlnt
