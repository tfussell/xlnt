// Copyright (c) 2014-2017 Thomas Fussell
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

#include <cstdint>
#include <string>

#include <xlnt/xlnt_config.hpp>

namespace xlnt {

enum class calendar;

/// <summary>
/// Describes the number formatting applied to text and numbers within a certain cell.
/// </summary>
class XLNT_API number_format
{
public:
    /// <summary>
    ///
    /// </summary>
    static const number_format general();

    /// <summary>
    ///
    /// </summary>
    static const number_format text();

    /// <summary>
    ///
    /// </summary>
    static const number_format number();

    /// <summary>
    ///
    /// </summary>
    static const number_format number_00();

    /// <summary>
    ///
    /// </summary>
    static const number_format number_comma_separated1();

    /// <summary>
    ///
    /// </summary>
    static const number_format percentage();

    /// <summary>
    ///
    /// </summary>
    static const number_format percentage_00();

    /// <summary>
    ///
    /// </summary>
    static const number_format date_yyyymmdd2();

    /// <summary>
    ///
    /// </summary>
    static const number_format date_yymmdd();

    /// <summary>
    ///
    /// </summary>
    static const number_format date_ddmmyyyy();

    /// <summary>
    ///
    /// </summary>
    static const number_format date_dmyslash();

    /// <summary>
    ///
    /// </summary>
    static const number_format date_dmyminus();

    /// <summary>
    ///
    /// </summary>
    static const number_format date_dmminus();

    /// <summary>
    ///
    /// </summary>
    static const number_format date_myminus();

    /// <summary>
    ///
    /// </summary>
    static const number_format date_xlsx14();

    /// <summary>
    ///
    /// </summary>
    static const number_format date_xlsx15();

    /// <summary>
    ///
    /// </summary>
    static const number_format date_xlsx16();

    /// <summary>
    ///
    /// </summary>
    static const number_format date_xlsx17();

    /// <summary>
    ///
    /// </summary>
    static const number_format date_xlsx22();

    /// <summary>
    ///
    /// </summary>
    static const number_format date_datetime();

    /// <summary>
    ///
    /// </summary>
    static const number_format date_time1();

    /// <summary>
    ///
    /// </summary>
    static const number_format date_time2();

    /// <summary>
    ///
    /// </summary>
    static const number_format date_time3();

    /// <summary>
    ///
    /// </summary>
    static const number_format date_time4();

    /// <summary>
    ///
    /// </summary>
    static const number_format date_time5();

    /// <summary>
    ///
    /// </summary>
    static const number_format date_time6();

    /// <summary>
    ///
    /// </summary>
    static bool is_builtin_format(std::size_t builtin_id);

    /// <summary>
    ///
    /// </summary>
    static const number_format &from_builtin_id(std::size_t builtin_id);

    /// <summary>
    ///
    /// </summary>
    number_format();

    /// <summary>
    ///
    /// </summary>
    number_format(std::size_t builtin_id);

    /// <summary>
    ///
    /// </summary>
    number_format(const std::string &code);

    /// <summary>
    ///
    /// </summary>
    number_format(const std::string &code, std::size_t custom_id);

    /// <summary>
    ///
    /// </summary>
    void format_string(const std::string &format_code);

    /// <summary>
    ///
    /// </summary>
    void format_string(const std::string &format_code, std::size_t custom_id);

    /// <summary>
    ///
    /// </summary>
    std::string format_string() const;

    /// <summary>
    ///
    /// </summary>
    bool has_id() const;

    /// <summary>
    ///
    /// </summary>
    void id(std::size_t id);

    /// <summary>
    ///
    /// </summary>
    std::size_t id() const;

    /// <summary>
    ///
    /// </summary>
    std::string format(const std::string &text) const;

    /// <summary>
    ///
    /// </summary>
    std::string format(long double number, calendar base_date) const;

    /// <summary>
    ///
    /// </summary>
    bool is_date_format() const;

    /// <summary>
    /// Returns true if left is exactly equal to right.
    /// </summary>
    XLNT_API friend bool operator==(const number_format &left, const number_format &right);

    /// <summary>
    /// Returns true if left is not exactly equal to right.
    /// </summary>
    XLNT_API friend bool operator!=(const number_format &left, const number_format &right)
    {
        return !(left == right);
    }

private:
    /// <summary>
    ///
    /// </summary>
    bool id_set_;

    /// <summary>
    ///
    /// </summary>
    std::size_t id_;

    /// <summary>
    ///
    /// </summary>
    std::string format_string_;
};

} // namespace xlnt
