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

namespace {

template <class T>
void hash_combine(std::size_t &seed, const T &v)
{
    std::hash<T> hasher;
    seed ^= hasher(v) + 0x9e3779b9 + (seed << 6) + (seed >> 2);
}
}

namespace xlnt {

std::string style::to_hash_string() const
{
    std::string hash_string("style");
    
    hash_string.append(name_);
    hash_string.append(std::to_string(format_id_));
    hash_string.append(std::to_string(builtin_id_));
    hash_string.append(std::to_string(hidden_ ? 0 : 1));

    return hash_string;
}

style::style() : name_("style"), format_id_(0), builtin_id_(0), hidden_(false)
{
}

style::style(const style &other) : name_(other.name_), format_id_(other.format_id_), builtin_id_(other.builtin_id_), hidden_(other.hidden_)
{
}

style &style::operator=(const xlnt::style &other)
{
    name_ = other.name_;
    builtin_id_ = other.builtin_id_;
    format_id_ = other.format_id_;
    hidden_ = other.hidden_;
    
    return *this;
}

std::string style::get_name() const
{
    return name_;
}

void style::set_name(const std::string &name)
{
    name_ = name;
}

std::size_t style::get_format_id() const
{
    return format_id_;
}

void style::set_format_id(std::size_t format_id)
{
    format_id_ = format_id;
}

std::size_t style::get_builtin_id() const
{
    return builtin_id_;
}

void style::set_builtin_id(std::size_t builtin_id)
{
    builtin_id_ = builtin_id;
}

void style::set_hidden(bool hidden)
{
    hidden_ = hidden;
}

bool style::get_hidden() const
{
    return hidden_;
}

} // namespace xlnt
