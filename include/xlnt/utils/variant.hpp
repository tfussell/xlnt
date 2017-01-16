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

#pragma once

#include <string>
#include <vector>

#include <xlnt/xlnt_config.hpp>

namespace xlnt {

struct datetime;

class XLNT_API variant
{
public:
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
        lpstr,
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

    variant();
    variant(const std::string &value);
    variant(const char *value);
    variant(int value);
    variant(bool value);
    variant(const datetime &value);
    variant(const std::initializer_list<int> &value);
    variant(const std::vector<int> &value);
    variant(const std::initializer_list<const char *> &value);
    variant(const std::vector<const char *> &value);
    variant(const std::initializer_list<std::string> &value);
    variant(const std::vector<std::string> &value);
    variant(const std::vector<variant> &value);

    bool is(type t) const;

    template<typename T>
    T get() const;

    type value_type() const;

private:
    type type_;
    std::vector<variant> vector_value_;
    std::int32_t i4_value_;
    std::string lpstr_value_;
};

} // namespace xlnt
