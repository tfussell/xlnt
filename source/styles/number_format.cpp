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

#include <algorithm>
#include <cctype>
#include <unordered_map>
#include <vector>

#include <xlnt/styles/number_format.hpp>
#include <xlnt/utils/datetime.hpp>
#include <xlnt/utils/exceptions.hpp>
#include <detail/number_formatter.hpp>

namespace {

const std::unordered_map<std::size_t, std::string> &builtin_formats()
{
    static const std::unordered_map<std::size_t, std::string> *formats =
        new std::unordered_map<std::size_t, std::string>(
            {{0, "General"}, {1, "0"}, {2, "0.00"}, {3, "#,##0"}, {4, "#,##0.00"}, {9, "0%"}, {10, "0.00%"},
                {11, "0.00E+00"}, {12, "# ?/?"}, {13, "# \?\?/??"}, // escape trigraph
                {14, "mm-dd-yy"}, {15, "d-mmm-yy"}, {16, "d-mmm"}, {17, "mmm-yy"}, {18, "h:mm AM/PM"},
                {19, "h:mm:ss AM/PM"}, {20, "h:mm"}, {21, "h:mm:ss"}, {22, "m/d/yy h:mm"}, {37, "#,##0 ;(#,##0)"},
                {38, "#,##0 ;[Red](#,##0)"}, {39, "#,##0.00;(#,##0.00)"}, {40, "#,##0.00;[Red](#,##0.00)"},

                // 41-44 aren't in the ECMA 376 v4 standard, but Libre Office uses them
                {41, "_(* #,##0_);_(* \\(#,##0\\);_(* \"-\"_);_(@_)"},
                {42, "_(\"$\"* #,##0_);_(\"$\"* \\(#,##0\\);_(\"$\"* \"-\"_);_(@_)"},
                {43, "_(* #,##0.00_);_(* \\(#,##0.00\\);_(* \"-\"??_);_(@_)"},
                {44, "_(\"$\"* #,##0.00_)_(\"$\"* \\(#,##0.00\\)_(\"$\"* \"-\"??_)_(@_)"},

                {45, "mm:ss"}, {46, "[h]:mm:ss"}, {47, "mmss.0"}, {48, "##0.0E+0"}, {49, "@"}});

    return *formats;
}

} // namespace

namespace xlnt {

const number_format number_format::general()
{
    static const number_format *format = new number_format(builtin_formats().at(0), 0);
    return *format;
}

const number_format number_format::text()
{
    static const number_format *format = new number_format(builtin_formats().at(49), 49);
    return *format;
}

const number_format number_format::number()
{
    static const number_format *format = new number_format(builtin_formats().at(1), 1);
    return *format;
}

const number_format number_format::number_00()
{
    static const number_format *format = new number_format(builtin_formats().at(2), 2);
    return *format;
}

const number_format number_format::number_comma_separated1()
{
    static const number_format *format = new number_format(builtin_formats().at(4), 4);
    return *format;
}

const number_format number_format::percentage()
{
    static const number_format *format = new number_format(builtin_formats().at(9), 9);
    return *format;
}

const number_format number_format::percentage_00()
{
    static const number_format *format = new number_format(builtin_formats().at(10), 10);
    return *format;
}

const number_format number_format::date_yyyymmdd2()
{
    static const number_format *format = new number_format("yyyy-mm-dd");
    return *format;
}

const number_format number_format::date_yymmdd()
{
    static const number_format *format = new number_format("yy-mm-dd");
    return *format;
}

const number_format number_format::date_ddmmyyyy()
{
    static const number_format *format = new number_format("dd/mm/yy");
    return *format;
}

const number_format number_format::date_dmyslash()
{
    static const number_format *format = new number_format("d/m/yy");
    return *format;
}

const number_format number_format::date_dmyminus()
{
    static const number_format *format = new number_format("d-m-yy");
    return *format;
}

const number_format number_format::date_dmminus()
{
    static const number_format *format = new number_format("d-m");
    return *format;
}

const number_format number_format::date_myminus()
{
    static const number_format *format = new number_format("m-yy");
    return *format;
}

const number_format number_format::date_xlsx14()
{
    static const number_format *format = new number_format(builtin_formats().at(14), 14);
    return *format;
}

const number_format number_format::date_xlsx15()
{
    static const number_format *format = new number_format(builtin_formats().at(15), 15);
    return *format;
}

const number_format number_format::date_xlsx16()
{
    static const number_format *format = new number_format(builtin_formats().at(16), 16);
    return *format;
}

const number_format number_format::date_xlsx17()
{
    static const number_format *format = new number_format(builtin_formats().at(17), 17);
    return *format;
}

const number_format number_format::date_xlsx22()
{
    static const number_format *format = new number_format(builtin_formats().at(22), 22);
    return *format;
}

const number_format number_format::date_datetime()
{
    static const number_format *format = new number_format("yyyy-mm-dd h:mm:ss");
    return *format;
}

const number_format number_format::date_time1()
{
    static const number_format *format = new number_format(builtin_formats().at(18), 18);
    return *format;
}

const number_format number_format::date_time2()
{
    static const number_format *format = new number_format(builtin_formats().at(19), 19);
    return *format;
}

const number_format number_format::date_time3()
{
    static const number_format *format = new number_format(builtin_formats().at(20), 20);
    return *format;
}

const number_format number_format::date_time4()
{
    static const number_format *format = new number_format(builtin_formats().at(21), 21);
    return *format;
}

const number_format number_format::date_time5()
{
    static const number_format *format = new number_format(builtin_formats().at(45), 45);
    return *format;
}

const number_format number_format::date_time6()
{
    static const number_format *format = new number_format(builtin_formats().at(21), 21);
    return *format;
}

number_format::number_format()
    : number_format(general())
{
}

number_format::number_format(std::size_t id)
    : number_format(from_builtin_id(id))
{
}

number_format::number_format(const std::string &format_string)
    : id_set_(false), id_(0)
{
    this->format_string(format_string);
}

number_format::number_format(const std::string &format_string, std::size_t id)
    : id_set_(false), id_(0)
{
    this->format_string(format_string, id);
}

number_format number_format::from_builtin_id(std::size_t builtin_id)
{
    if (builtin_formats().find(builtin_id) == builtin_formats().end())
    {
        throw invalid_parameter(); //("unknown id: " + std::to_string(builtin_id));
    }

    auto format_string = builtin_formats().at(builtin_id);
    return number_format(format_string, builtin_id);
}

std::string number_format::format_string() const
{
    return format_string_;
}

void number_format::format_string(const std::string &format_string)
{
    format_string_ = format_string;
    id_ = 0;
    id_set_ = false;

    for (const auto &pair : builtin_formats())
    {
        if (pair.second == format_string)
        {
            id_ = pair.first;
            id_set_ = true;
            break;
        }
    }
}

void number_format::format_string(const std::string &format_string, std::size_t id)
{
    format_string_ = format_string;
    id_ = id;
    id_set_ = true;
}

bool number_format::has_id() const
{
    return id_set_;
}

void number_format::id(std::size_t id)
{
    id_ = id;
    id_set_ = true;
}

std::size_t number_format::id() const
{
    if (!id_set_)
    {
        throw invalid_attribute();
    }

    return id_;
}

bool number_format::is_date_format() const
{
    detail::number_format_parser p(format_string_);
    p.parse();
    auto parsed = p.result();

    bool any_datetime = false;
    bool any_timedelta = false;

    for (const auto &section : parsed)
    {
        if (section.is_datetime)
        {
            any_datetime = true;
        }

        if (section.is_timedelta)
        {
            any_timedelta = true;
        }
    }

    return any_datetime && !any_timedelta;
}

std::string number_format::format(const std::string &text) const
{
    return detail::number_formatter(format_string_, calendar::windows_1900).format_text(text);
}

std::string number_format::format(long double number, calendar base_date) const
{
    return detail::number_formatter(format_string_, base_date).format_number(number);
}

XLNT_API bool operator==(const number_format &left, const number_format &right)
{
    return left.format_string_ == right.format_string_;
}

} // namespace xlnt
