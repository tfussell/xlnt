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
#include <xlnt/packaging/override_type.hpp>

namespace xlnt {

override_type::override_type()
{
}

override_type::override_type(const std::string &part_name, const std::string &content_type)
    : part_name_(part_name), content_type_(content_type)
{
}

override_type::override_type(const override_type &other)
    : part_name_(other.part_name_), content_type_(other.content_type_)
{
}

override_type &override_type::operator=(const override_type &other)
{
    part_name_ = other.part_name_;
    content_type_ = other.content_type_;

    return *this;
}

std::string override_type::get_part_name() const
{
    return part_name_;
}

std::string override_type::get_content_type() const
{
    return content_type_;
}

} // namespace xlnt
