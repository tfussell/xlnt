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

alignment &format::alignment()
{
    return d_->parent->alignments.at(d_->alignment_id);
}

const alignment &format::alignment() const
{
    return d_->parent->alignments.at(d_->alignment_id);
}

void format::alignment(const xlnt::alignment &new_alignment, bool applied, bool update)
{
    auto alignment_id = d_->parent->add_alignment(new_alignment);

    if (update)
    {
        d_->alignment_id = alignment_id;
        d_->alignment_applied = applied;
    }
    else
    {
        auto impl_copy = *d_;

        impl_copy.alignment_id = alignment_id;
        impl_copy.alignment_applied = applied;

        auto match = std::find(d_->parent->format_impls.begin(), d_->parent->format_impls.end(), impl_copy);

        if (match != d_->parent->format_impls.end())
        {
            d_ = &*match;
        }
        else
        {
            d_->parent->format_impls.push_back(impl_copy);
            d_ = &d_->parent->format_impls.back();
        }
    }
}

xlnt::border &format::border()
{
    const auto &self = *this;
    return const_cast<class border &>(self.border());
}

const xlnt::border &format::border() const
{
    if (!d_->border_id)
    {
        throw xlnt::exception("no border");
    }

	return d_->parent->borders.at(d_->border_id.get());
}

void format::border(const xlnt::border &new_border, bool applied, bool update)
{
    auto border_id = d_->parent->add_border(new_border);

    if (update)
    {
        d_->border_id = border_id;
        d_->border_applied = applied;
    }
    else
    {
        auto impl_copy = *d_;

        impl_copy.border_id = border_id;
        impl_copy.border_applied = applied;

        auto match = std::find(d_->parent->format_impls.begin(), d_->parent->format_impls.end(), impl_copy);

        if (match != d_->parent->format_impls.end())
        {
            d_ = &*match;
        }
        else
        {
            d_->parent->format_impls.push_back(impl_copy);
            d_ = &d_->parent->format_impls.back();
        }
    }
}

xlnt::fill &format::fill()
{
    const auto &self = *this;
    return const_cast<class fill &>(self.fill());
}

const xlnt::fill &format::fill() const
{
    if (!d_->fill_id)
    {
        throw xlnt::exception("no border");
    }

	return d_->parent->fills.at(d_->fill_id.get());
}

void format::fill(const xlnt::fill &new_fill, bool applied, bool update)
{
    auto fill_id = d_->parent->add_fill(new_fill);

    if (update)
    {
        d_->fill_id = fill_id;
        d_->fill_applied = applied;
    }
    else
    {
        auto impl_copy = *d_;

        impl_copy.fill_id = fill_id;
        impl_copy.fill_applied = applied;

        auto match = std::find(d_->parent->format_impls.begin(), d_->parent->format_impls.end(), impl_copy);

        if (match != d_->parent->format_impls.end())
        {
            d_ = &*match;
        }
        else
        {
            d_->parent->format_impls.push_back(impl_copy);
            d_ = &d_->parent->format_impls.back();
        }
    }
}

xlnt::font &format::font()
{
    const auto &self = *this;
    return const_cast<class font &>(self.font());
}

const xlnt::font &format::font() const
{
    if (!d_->font_id)
    {
        throw xlnt::exception("no border");
    }

	return d_->parent->fonts.at(d_->font_id.get());
}

void format::font(const xlnt::font &new_font, bool applied, bool update)
{
    auto font_id = d_->parent->add_font(new_font);

    if (update)
    {
        d_->font_id = font_id;
        d_->font_applied = applied;
    }
    else
    {
        auto impl_copy = *d_;

        impl_copy.font_id = font_id;
        impl_copy.font_applied = applied;

        auto match = std::find(d_->parent->format_impls.begin(), d_->parent->format_impls.end(), impl_copy);

        if (match != d_->parent->format_impls.end())
        {
            d_ = &*match;
        }
        else
        {
            d_->parent->format_impls.push_back(impl_copy);
            d_ = &d_->parent->format_impls.back();
        }
    }
}

xlnt::number_format &format::number_format()
{
    const auto &self = *this;
    return const_cast<class number_format &>(self.number_format());
}

const xlnt::number_format &format::number_format() const
{
    if (!d_->number_format_id)
    {
        throw xlnt::exception("no number format");
    }

	return d_->parent->number_formats.at(d_->number_format_id.get());
}

void format::number_format(const xlnt::number_format &new_number_format, bool applied, bool update)
{
	auto copy = new_number_format;

	if (!copy.has_id())
	{
		copy.set_id(d_->parent->next_custom_number_format_id());
		d_->parent->number_formats[copy.get_id()] = copy;
	}

    auto number_format_id = copy.get_id();

    if (update)
    {
        d_->number_format_id = number_format_id;
        d_->number_format_applied = applied;
    }
    else
    {
        auto impl_copy = *d_;

        impl_copy.number_format_id = number_format_id;
        impl_copy.number_format_applied = applied;

        auto match = std::find(d_->parent->format_impls.begin(), d_->parent->format_impls.end(), impl_copy);

        if (match != d_->parent->format_impls.end())
        {
            d_ = &*match;
        }
        else
        {
            d_->parent->format_impls.push_back(impl_copy);
            d_ = &d_->parent->format_impls.back();
        }
    }
}

protection &format::protection()
{
    return d_->parent->protections.at(d_->protection_id);
}

const protection &format::protection() const
{
    return d_->parent->protections.at(d_->protection_id);
}

void format::protection(const xlnt::protection &new_protection, bool applied, bool update)
{
    auto protection_id = d_->parent->add_protection(new_protection);

    if (update)
    {
        d_->protection_id = protection_id;
        d_->protection_applied = applied;
    }
    else
    {
        auto impl_copy = *d_;

        impl_copy.protection_id = protection_id;
        impl_copy.protection_applied = applied;

        auto match = std::find(d_->parent->format_impls.begin(), d_->parent->format_impls.end(), impl_copy);

        if (match != d_->parent->format_impls.end())
        {
            d_ = &*match;
        }
        else
        {
            d_->parent->format_impls.push_back(impl_copy);
            d_ = &d_->parent->format_impls.back();
        }
    }
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



XLNT_API bool operator==(const format &left, const format &right)
{
    if (left.alignment_applied() != right.alignment_applied()) return false;
    if (left.alignment_applied() && left.alignment() != right.alignment()) return false;

    if (left.border_applied() != right.border_applied()) return false;
    if (left.border_applied() && left.border() != right.border()) return false;

    if (left.fill_applied() != right.fill_applied()) return false;
    if (left.fill_applied() && left.fill() != right.fill()) return false;

    if (left.font_applied() != right.font_applied()) return false;
    if (left.font_applied() && left.font() != right.font()) return false;

    if (left.number_format_applied() != right.number_format_applied()) return false;
    if (left.number_format_applied() && left.number_format() != right.number_format()) return false;

    if (left.protection_applied() != right.protection_applied()) return false;
    if (left.protection_applied() && left.protection() != right.protection()) return false;
    
    if (left.d_->style.is_set() != right.d_->style.is_set()) return false;
    if (left.d_->style.is_set() && left.d_->style.get() != right.d_->style.get()) return false;

    return true;
}

} // namespace xlnt
