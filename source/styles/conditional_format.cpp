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

#include <detail/implementations/conditional_format_impl.hpp>
#include <detail/implementations/stylesheet.hpp>
#include <xlnt/styles/border.hpp>
#include <xlnt/styles/fill.hpp>
#include <xlnt/styles/font.hpp>
#include <xlnt/styles/conditional_format.hpp>

namespace xlnt {

condition condition::text_starts_with(const std::string &text)
{
	condition c;
	c.type_ = type::contains_text;
	c.operator_ = condition_operator::starts_with;
	c.text_comparand_ = text;
	return c;
}

condition condition::text_ends_with(const std::string &text)
{
	condition c;
	c.type_ = type::contains_text;
	c.operator_ = condition_operator::ends_with;
	c.text_comparand_ = text;
	return c;
}

condition condition::text_contains(const std::string &text)
{
	condition c;
	c.type_ = type::contains_text;
	c.operator_ = condition_operator::contains;
	c.text_comparand_ = text;
	return c;
}

condition condition::text_does_not_contain(const std::string &text)
{
	condition c;
	c.type_ = type::contains_text;
	c.operator_ = condition_operator::does_not_contain;
	c.text_comparand_ = text;
	return c;
}

conditional_format::conditional_format(detail::conditional_format_impl *d) : d_(d)
{
}

bool conditional_format::operator==(const conditional_format &other) const
{
	return d_ == other.d_;
}

bool conditional_format::operator!=(const conditional_format &other) const
{
	return !(*this == other);
}

xlnt::border conditional_format::border() const
{
    return d_->parent->borders.at(d_->border_id.get());
}

conditional_format conditional_format::border(const xlnt::border &new_border)
{
    d_->border_id = d_->parent->find_or_add(d_->parent->borders, new_border);
	return *this;
}

xlnt::fill conditional_format::fill() const
{
    return d_->parent->fills.at(d_->fill_id.get());
}

conditional_format conditional_format::fill(const xlnt::fill &new_fill)
{
    d_->fill_id = d_->parent->find_or_add(d_->parent->fills, new_fill);
	return *this;
}

xlnt::font conditional_format::font() const
{
    return d_->parent->fonts.at(d_->font_id.get());
}

conditional_format conditional_format::font(const xlnt::font &new_font)
{
    d_->font_id = d_->parent->find_or_add(d_->parent->fonts, new_font);
	return *this;
}

} // namespace xlnt
