// Copyright (c) 2014 Thomas Fussell
// Copyright (c) 2010-2014 openpyxl
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

#include <memory>
#include <string>
#include <unordered_map>

#include "../styles/style.hpp"
#include "../common/types.hpp"

namespace xlnt {

struct date;
struct datetime;
struct time;

class cell_value
{
public:
    enum class type
    {
        null,
        numeric,
        string,
        formula,
        boolean,
        error
    };

    std::string to_string() const;

    type get_data_type() const;

    void set_null();

    cell_value &operator=(const cell_value &rhs);
    cell_value &operator=(bool value);
    cell_value &operator=(int value);
    cell_value &operator=(double value);
    cell_value &operator=(long int value);
    cell_value &operator=(long long value);
    cell_value &operator=(long double value);
    cell_value &operator=(const std::string &value);
    cell_value &operator=(const char *value);
    cell_value &operator=(const date &value);
    cell_value &operator=(const time &value);
    cell_value &operator=(const datetime &value);

    bool operator==(const cell_value &comparand) const;
    bool operator==(std::nullptr_t) const;
    bool operator==(bool comparand) const;
    bool operator==(int comparand) const;
    bool operator==(double comparand) const;
    bool operator==(const std::string &comparand) const;
    bool operator==(const char *comparand) const;
    bool operator==(const date &comparand) const;
    bool operator==(const time &comparand) const;
    bool operator==(const datetime &comparand) const;

    friend bool operator==(std::nullptr_t, const cell_value &cell);
    friend bool operator==(bool comparand, const cell_value &cell);
    friend bool operator==(int comparand, const cell_value &cell);
    friend bool operator==(double comparand, const cell_value &cell);
    friend bool operator==(const std::string &comparand, const cell_value &cell);
    friend bool operator==(const char *comparand, const cell_value &cell);
    friend bool operator==(const date &comparand, const cell_value &cell);
    friend bool operator==(const time &comparand, const cell_value &cell);
    friend bool operator==(const datetime &comparand, const cell_value &cell);
};

inline std::ostream &operator<<(std::ostream &stream, const xlnt::cell_value &value)
{
    return stream << value.to_string();
}

} // namespace xlnt
