// Copyright (c) 2017-2018 Thomas Fussell
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

#pragma once

#include <string>
#include <vector>

#include <xlnt/xlnt_config.hpp>

namespace xlnt {

struct datetime;

/// <summary>
/// Represents an object that can have variable type.
/// </summary>
class XLNT_API variant
{
public:
    // TODO: implement remaining types?

    /// <summary>
    /// The possible types a variant can hold.
    /// </summary>
    enum class type
    {
        vector,
        //array,
        //blob,
        //oblob,
        //empty,
        null,
        //i1,
        //i2,
        i4,
        //i8,
        //integer,
        //ui1,
        //ui2,
        //ui4,
        //ui8,
        //uint,
        //r4,
        //r8,
        //decimal,
        lpstr, // TODO: how does this differ from lpwstr?
        //lpwstr,
        //bstr,
        date,
        //filetime,
        boolean,
        //cy,
        //error,
        //stream,
        //ostream,
        //storage,
        //ostorage,
        //vstream,
        //clsid
    };

    /// <summary>
    /// Default constructor. Creates a null-type variant.
    /// </summary>
    variant();

    /// <summary>
    /// Creates a string-type variant with the given value.
    /// </summary>
    variant(const std::string &value);

    /// <summary>
    /// Creates a string-type variant with the given value.
    /// </summary>
    variant(const char *value);

    /// <summary>
    /// Creates a i4-type variant with the given value.
    /// </summary>
    variant(std::int32_t value);

    /// <summary>
    /// Creates a bool-type variant with the given value.
    /// </summary>
    variant(bool value);

    /// <summary>
    /// Creates a date-type variant with the given value.
    /// </summary>
    variant(const datetime &value);

    /// <summary>
    /// Creates a vector_i4-type variant with the given value.
    /// </summary>
    variant(const std::initializer_list<std::int32_t> &value);

    /// <summary>
    /// Creates a vector_i4-type variant with the given value.
    /// </summary>
    variant(const std::vector<std::int32_t> &value);

    /// <summary>
    /// Creates a vector_string-type variant with the given value.
    /// </summary>
    variant(const std::initializer_list<const char *> &value);

    /// <summary>
    /// Creates a vector_string-type variant with the given value.
    /// </summary>
    variant(const std::vector<const char *> &value);

    /// <summary>
    /// Creates a vector_string-type variant with the given value.
    /// </summary>
    variant(const std::initializer_list<std::string> &value);

    /// <summary>
    /// Creates a vector_string-type variant with the given value.
    /// </summary>
    variant(const std::vector<std::string> &value);

    /// <summary>
    /// Creates a vector_variant-type variant with the given value.
    /// </summary>
    variant(const std::vector<variant> &value);

    /// <summary>
    /// Returns true if this variant is of type t.
    /// </summary>
    bool is(type t) const;

    /// <summary>
    /// Returns the value of this variant as type T. An exception will
    /// be thrown if the types are not convertible.
    /// </summary>
    template<typename T>
    T get() const;

    /// <summary>
    /// Returns the type of this variant.
    /// </summary>
    type value_type() const;

    bool operator==(const variant &rhs) const;

private:
    type type_;
    std::vector<variant> vector_value_;
    std::int32_t i4_value_;
    std::string lpstr_value_;
};

template<>
bool variant::get() const;

template<>
std::int32_t variant::get() const;

template<>
std::string variant::get() const;

template<>
datetime variant::get() const;

template<>
std::vector<std::int32_t> variant::get() const;

template<>
std::vector<std::string> variant::get() const;

template<>
std::vector<variant> variant::get() const;

} // namespace xlnt
