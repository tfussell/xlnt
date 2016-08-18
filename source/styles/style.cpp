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

#include <xlnt/styles/alignment.hpp>
#include <xlnt/styles/border.hpp>
#include <xlnt/styles/fill.hpp>
#include <xlnt/styles/font.hpp>
#include <xlnt/styles/number_format.hpp>
#include <xlnt/styles/protection.hpp>
#include <detail/style_impl.hpp> // include order is important here
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

bool style::operator==(const style &other)
{
	return d_ == other.d_;
}

} // namespace xlnt
