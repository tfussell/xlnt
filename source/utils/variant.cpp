// Copyright (c) 2017 Thomas Fussell
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

#include <xlnt/utils/variant.hpp>
#include <xlnt/utils/datetime.hpp>

namespace xlnt {

variant::variant()
{

}

variant::variant(const std::string &value)
    : type_(type::lpstr),
    lpstr_value_(value)
{
}

variant::variant(const char *value) : variant(std::string(value))
{
}

variant::variant(int value)
    : type_(type::i4),
    i4_value_(value)
{

}

variant::variant(bool value)
    : type_(type::boolean),
    i4_value_(value ? 1 : 0)
{

}

variant::variant(const datetime &value)
    : type_(type::date),
    lpstr_value_(value.to_iso_string())
{

}

variant::variant(const std::initializer_list<int> &value)
    : type_(type::vector)
{
    for (const auto &v : value)
    {
        vector_value_.emplace_back(v);
    }
}

variant::variant(const std::vector<int> &value)
    : type_(type::vector)
{
    for (const auto &v : value)
    {
        vector_value_.emplace_back(v);
    }
}

variant::variant(const std::initializer_list<const char *> &value)
    : type_(type::vector)
{
    for (const auto &v : value)
    {
        vector_value_.emplace_back(v);
    }
}

variant::variant(const std::vector<const char *> &value)
    : type_(type::vector)
{
    for (const auto &v : value)
    {
        vector_value_.emplace_back(v);
    }
}

variant::variant(const std::initializer_list<std::string> &value)
    : type_(type::vector)
{
    for (const auto &v : value)
    {
        vector_value_.emplace_back(v);
    }
}

variant::variant(const std::vector<std::string> &value)
    : type_(type::vector)
{
    for (const auto &v : value)
    {
        vector_value_.emplace_back(v);
    }
}

variant::variant(const std::vector<variant> &value)
    : type_(type::vector)
{
    for (const auto &v : value)
    {
        vector_value_.emplace_back(v);
    }
}

bool variant::is(type t) const
{
    return type_ == t;
}

template<>
XLNT_API std::string variant::get() const
{
    return lpstr_value_;
}

template<>
XLNT_API std::vector<variant> variant::get() const
{
    return vector_value_;
}

template<>
XLNT_API bool variant::get() const
{
    return i4_value_ != 0;
}

template<>
XLNT_API std::int32_t variant::get() const
{
    return i4_value_;
}

template<>
XLNT_API datetime variant::get() const
{
    return datetime::from_iso_string(lpstr_value_);
}

variant::type variant::value_type() const
{
    return type_;
}

} // namespace xlnt
