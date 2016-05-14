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

#include <xlnt/styles/style.hpp>

namespace xlnt {

style::style()
    : formattable(),
      hidden_(false),
      builtin_id_(0)
{
}

style::style(const style &other)
    : formattable(other),
      name_(other.name_),
      hidden_(other.hidden_),
      builtin_id_(other.builtin_id_)
{
}

style &style::operator=(const style &other)
{
    formattable::operator=(other);
    
    name_ = other.name_;
    hidden_ = other.hidden_;
    builtin_id_ = other.builtin_id_;
    
    return *this;
}

bool style::get_hidden() const
{
    return hidden_;
}

void style::set_hidden(bool value)
{
    hidden_ = value;
}

std::size_t style::get_builtin_id() const
{
    return builtin_id_;
}

void style::set_builtin_id(std::size_t builtin_id)
{
    builtin_id_ = builtin_id;
}

std::string style::get_name() const
{
    return name_;
}

void style::set_name(const std::string &name)
{
    name_ = name;
}

std::string style::to_hash_string() const
{
    return formattable::to_hash_string() + name_;
}

} // namespace xlnt
