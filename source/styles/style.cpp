// Copyright (c) 2014-2018 Thomas Fussell
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

#include <detail/implementations/style_impl.hpp>
#include <detail/implementations/stylesheet.hpp>
#include <xlnt/styles/alignment.hpp>
#include <xlnt/styles/border.hpp>
#include <xlnt/styles/fill.hpp>
#include <xlnt/styles/font.hpp>
#include <xlnt/styles/number_format.hpp>
#include <xlnt/styles/protection.hpp>
#include <xlnt/styles/style.hpp>

namespace {

std::vector<xlnt::number_format>::iterator find_number_format(
	std::vector<xlnt::number_format> &number_formats, std::size_t id)
{
	return std::find_if(number_formats.begin(), number_formats.end(),
		[=](const xlnt::number_format &nf) { return nf.id() == id; });
}

} // namespace

namespace xlnt {

style::style(detail::style_impl *d)
    : d_(d)
{
}

bool style::hidden() const
{
    return d_->hidden_style;
}

style style::hidden(bool value)
{
    d_->hidden_style = value;
    return style(d_);
}

std::size_t style::builtin_id() const
{
    return d_->builtin_id.get();
}

bool style::builtin() const
{
    return d_->builtin_id.is_set();
}

std::string style::name() const
{
    return d_->name;
}

style style::name(const std::string &name)
{
    d_->name = name;
    return *this;
}

bool style::custom_builtin() const
{
	return d_->builtin_id.is_set() && d_->custom_builtin;
}

bool style::operator==(const style &other) const
{
    return name() == other.name();
}

bool style::operator!=(const style &other) const
{
    return !operator==(other);
}

xlnt::alignment style::alignment() const
{
    return d_->parent->alignments.at(d_->alignment_id.get());
}

style style::alignment(const xlnt::alignment &new_alignment, optional<bool> applied)
{
    d_->alignment_id = d_->parent->find_or_add(d_->parent->alignments, new_alignment);
    d_->alignment_applied = applied;

	return *this;
}

xlnt::border style::border() const
{
    return d_->parent->borders.at(d_->border_id.get());
}

style style::border(const xlnt::border &new_border, optional<bool> applied)
{
    d_->border_id = d_->parent->find_or_add(d_->parent->borders, new_border);
    d_->border_applied = applied;

	return *this;
}

xlnt::fill style::fill() const
{
    return d_->parent->fills.at(d_->fill_id.get());
}

style style::fill(const xlnt::fill &new_fill, optional<bool> applied)
{
    d_->fill_id = d_->parent->find_or_add(d_->parent->fills, new_fill);
    d_->fill_applied = applied;

	return *this;
}

xlnt::font style::font() const
{
    return d_->parent->fonts.at(d_->font_id.get());
}

style style::font(const xlnt::font &new_font, optional<bool> applied)
{
    d_->font_id = d_->parent->find_or_add(d_->parent->fonts, new_font);
    d_->font_applied = applied;

	return *this;
}

xlnt::number_format style::number_format() const
{
	auto match = find_number_format(d_->parent->number_formats, 
		d_->number_format_id.get());

	if (match == d_->parent->number_formats.end())
	{
		throw invalid_attribute();
	}

	return *match;
}

style style::number_format(const xlnt::number_format &new_number_format, optional<bool> applied)
{
    auto copy = new_number_format;

    if (!copy.has_id())
    {
        copy.id(d_->parent->next_custom_number_format_id());
        d_->parent->number_formats.push_back(copy);
    }
	else if (find_number_format(d_->parent->number_formats, copy.id()) 
		== d_->parent->number_formats.end())
	{
        d_->parent->number_formats.push_back(copy);
    }

    d_->number_format_id = copy.id();
    d_->number_format_applied = applied;

	return *this;
}

xlnt::protection style::protection() const
{
    return d_->parent->protections.at(d_->protection_id.get());
}

style style::protection(const xlnt::protection &new_protection, optional<bool> applied)
{
    d_->protection_id = d_->parent->find_or_add(d_->parent->protections, new_protection);
    d_->protection_applied = applied;

    return *this;
}

bool style::alignment_applied() const
{
    return d_->alignment_applied.is_set()
        ? d_->alignment_applied.get()
        : d_->alignment_id.is_set();
}

bool style::border_applied() const
{
    return d_->border_applied.is_set()
        ? d_->border_applied.get()
        : d_->border_id.is_set();
}

bool style::fill_applied() const
{
    return d_->fill_applied.is_set()
        ? d_->fill_applied.get()
        : d_->fill_id.is_set();
}

bool style::font_applied() const
{
    return d_->font_applied.is_set()
        ? d_->font_applied.get()
        : d_->font_id.is_set();
}

bool style::number_format_applied() const
{
    return d_->number_format_applied.is_set()
        ? d_->number_format_applied.get()
        : d_->number_format_id.is_set();
}

bool style::protection_applied() const
{
    return d_->protection_applied.is_set()
        ? d_->protection_applied.get()
        : d_->protection_id.is_set();
}

bool style::pivot_button() const
{
    return d_->pivot_button_;
}

void style::pivot_button(bool show)
{
    d_->pivot_button_ = show;
}

bool style::quote_prefix() const
{
    return d_->quote_prefix_;
}

void style::quote_prefix(bool quote)
{
    d_->quote_prefix_ = quote;
}

} // namespace xlnt
