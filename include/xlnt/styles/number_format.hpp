// Copyright (c) 2015 Thomas Fussell
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
#pragma once

#include <string>
#include <unordered_map>
#include <utility>

namespace xlnt {

class number_format
{
public:
    static const number_format general();
    static const number_format text();
    static const number_format number();
    static const number_format number_00();
    static const number_format number_comma_separated1();
    static const number_format number_comma_separated2();
    static const number_format percentage();
    static const number_format percentage_00();
    static const number_format date_yyyymmdd2();
    static const number_format date_yyyymmdd();
    static const number_format date_ddmmyyyy();
    static const number_format date_dmyslash();
    static const number_format date_dmyminus();
    static const number_format date_dmminus();
    static const number_format date_myminus();
    static const number_format date_xlsx14();
    static const number_format date_xlsx15();
    static const number_format date_xlsx16();
    static const number_format date_xlsx17();
    static const number_format date_xlsx22();
    static const number_format date_datetime();
    static const number_format date_time1();
    static const number_format date_time2();
    static const number_format date_time3();
    static const number_format date_time4();
    static const number_format date_time5();
    static const number_format date_time6();
    static const number_format date_time7();
    static const number_format date_time8();
    static const number_format date_timedelta();
    static const number_format date_yyyymmddslash();
    static const number_format currency_usd_simple();
    static const number_format currency_usd();
    static const number_format currency_eur_simple();

    static number_format from_builtin_id(int builtin_id);
    
    number_format();
    number_format(int builtin_id);
    number_format(const std::string &code, int custom_id = -1);
    
    void set_format_string(const std::string &format_code, int id = -1);
    std::string get_format_string() const;

    void set_id(int id) { id_ = id; }
    int get_id() const { return id_; }
    
    std::size_t hash() const;
    
    bool operator==(const number_format &other) const
    {
        return hash() == other.hash();
    }

private:
    int id_;
    std::string format_string_;
};

} // namespace xlnt
