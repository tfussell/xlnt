// Copyright (c) 2014-2017 Thomas Fussell
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

#include <xlnt/styles/alignment.hpp>
#include <xlnt/styles/border.hpp>
#include <xlnt/styles/fill.hpp>
#include <xlnt/styles/font.hpp>
#include <xlnt/styles/number_format.hpp>
#include <xlnt/styles/protection.hpp>
#include <xlnt/styles/style.hpp>
#include <detail/stylesheet.hpp>

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

optional<std::size_t> style::builtin_id() const
{
    return d_->builtin_id;
}

style style::builtin_id(std::size_t builtin_id)
{
    d_->builtin_id = builtin_id;
    return *this;
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

optional<bool> style::custom() const
{
    return d_->custom_builtin;
}

style style::custom(bool value)
{
    d_->custom_builtin = value;
    return *this;
}

bool style::operator==(const style &other) const
{
    return name() == other.name();
}

xlnt::alignment &style::alignment()
{
    return d_->parent->alignments.at(d_->alignment_id.get());
}

const xlnt::alignment &style::alignment() const
{
    return d_->parent->alignments.at(d_->alignment_id.get());
}

style style::alignment(const xlnt::alignment &new_alignment, bool applied)
{
    d_->alignment_id = d_->parent->find_or_add(d_->parent->alignments, new_alignment);
    d_->alignment_applied = applied;
    return style(d_);
}

xlnt::border &style::border()
{
    return d_->parent->borders.at(d_->border_id.get());
}

const xlnt::border &style::border() const
{
    return d_->parent->borders.at(d_->border_id.get());
}

style style::border(const xlnt::border &new_border, bool applied)
{
    d_->border_id = d_->parent->find_or_add(d_->parent->borders, new_border);
    d_->border_applied = applied;
    return style(d_);
}

xlnt::fill &style::fill()
{
    return d_->parent->fills.at(d_->fill_id.get());
}

const xlnt::fill &style::fill() const
{
    return d_->parent->fills.at(d_->fill_id.get());
}

style style::fill(const xlnt::fill &new_fill, bool applied)
{
    d_->fill_id = d_->parent->find_or_add(d_->parent->fills, new_fill);
    d_->fill_applied = applied;
    return style(d_);
}

xlnt::font &style::font()
{
    return d_->parent->fonts.at(d_->font_id.get());
}

const xlnt::font &style::font() const
{
    return d_->parent->fonts.at(d_->font_id.get());
}

style style::font(const xlnt::font &new_font, bool applied)
{
    d_->font_id = d_->parent->find_or_add(d_->parent->fonts, new_font);
    d_->font_applied = applied;
    return style(d_);
}

xlnt::number_format &style::number_format()
{
    auto tarid = d_->number_format_id.get();
    return *std::find_if(d_->parent->number_formats.begin(), d_->parent->number_formats.end(),
        [=](const class number_format &nf) { return nf.id() == tarid; });
}

const xlnt::number_format &style::number_format() const
{
    auto tarid = d_->number_format_id.get();
    return *std::find_if(d_->parent->number_formats.begin(), d_->parent->number_formats.end(),
        [=](const class number_format &nf) { return nf.id() == tarid; });
}

style style::number_format(const xlnt::number_format &new_number_format, bool applied)
{
    auto copy = new_number_format;

    if (!copy.has_id())
    {
        copy.id(d_->parent->next_custom_number_format_id());
        d_->parent->number_formats.push_back(copy);
    }
    else if (std::find_if(d_->parent->number_formats.begin(), d_->parent->number_formats.end(),
                 [&copy](const class number_format &nf) { return nf.id() == copy.id(); })
        == d_->parent->number_formats.end())
    {
        d_->parent->number_formats.push_back(copy);
    }

    d_->number_format_id = copy.id();
    d_->number_format_applied = applied;

    return style(d_);
}

xlnt::protection &style::protection()
{
    return d_->parent->protections.at(d_->protection_id.get());
}

const xlnt::protection &style::protection() const
{
    return d_->parent->protections.at(d_->protection_id.get());
}

style style::protection(const xlnt::protection &new_protection, bool applied)
{
    d_->protection_id = d_->parent->find_or_add(d_->parent->protections, new_protection);
    d_->protection_applied = applied;
    return style(d_);
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

} // namespace xlnt
