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
#include <xlnt/styles/named_style.hpp>

namespace xlnt {

named_style::named_style()
{
}

named_style::named_style(const named_style &other)
{
}

bool named_style::get_hidden() const
{
    return hidden_;
}

void named_style::set_hidden(bool value)
{
    hidden_ = value;
}

std::size_t named_style::get_builtin_id() const
{
    return builtin_id_;
}

void named_style::set_builtin_id(std::size_t builtin_id)
{
    builtin_id_ = builtin_id;
}

std::string named_style::get_name() const
{
    return name_;
}

void named_style::set_name(const std::string &name)
{
    name_ = name;
}

std::string named_style::to_hash_string() const
{
    return common_style::to_hash_string() + name_;
}

} // namespace xlnt
