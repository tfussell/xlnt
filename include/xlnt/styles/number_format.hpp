// Copyright (c) 2014-2021 Thomas Fussell
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
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE
//
// @license: http://www.opensource.org/licenses/mit-license.php
// @author: see AUTHORS file

#pragma once

#include <cstdint>
#include <string>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/utils/optional.hpp>

namespace xlnt {

enum class calendar;

/// <summary>
/// Describes the number formatting applied to text and numbers within a certain cell.
/// </summary>
class XLNT_API number_format
{
public:
    /// <summary>
    /// Number format "General"
    /// </summary>
    static const number_format general();

    /// <summary>
    /// Number format "@"
    /// </summary>
    static const number_format text();

    /// <summary>
    /// Number format "0"
    /// </summary>
    static const number_format number();

    /// <summary>
    /// Number format "00"
    /// </summary>
    static const number_format number_00();

    /// <summary>
    /// Number format "#,##0.00"
    /// </summary>
    static const number_format number_comma_separated1();

    /// <summary>
    /// Number format "0%"
    /// </summary>
    static const number_format percentage();

    /// <summary>
    /// Number format "0.00%"
    /// </summary>
    static const number_format percentage_00();

    /// <summary>
    /// Number format "yyyy-mm-dd"
    /// </summary>
    static const number_format date_yyyymmdd2();

    /// <summary>
    /// Number format "yy-mm-dd"
    /// </summary>
    static const number_format date_yymmdd();

    /// <summary>
    /// Number format "dd/mm/yy"
    /// </summary>
    static const number_format date_ddmmyyyy();

    /// <summary>
    /// Number format "d/m/yy"
    /// </summary>
    static const number_format date_dmyslash();

    /// <summary>
    /// Number format "d-m-yy"
    /// </summary>
    static const number_format date_dmyminus();

    /// <summary>
    /// Number format "d-m"
    /// </summary>
    static const number_format date_dmminus();

    /// <summary>
    /// Number format "m-yy"
    /// </summary>
    static const number_format date_myminus();

    /// <summary>
    /// Number format "mm-dd-yy"
    /// </summary>
    static const number_format date_xlsx14();

    /// <summary>
    /// Number format "d-mmm-yy"
    /// </summary>
    static const number_format date_xlsx15();

    /// <summary>
    /// Number format "d-mmm"
    /// </summary>
    static const number_format date_xlsx16();

    /// <summary>
    /// Number format "mmm-yy"
    /// </summary>
    static const number_format date_xlsx17();

    /// <summary>
    /// Number format "m/d/yy h:mm"
    /// </summary>
    static const number_format date_xlsx22();

    /// <summary>
    /// Number format "yyyy-mm-dd h:mm:ss"
    /// </summary>
    static const number_format date_datetime();

    /// <summary>
    /// Number format "h:mm AM/PM"
    /// </summary>
    static const number_format date_time1();

    /// <summary>
    /// Number format "h:mm:ss AM/PM"
    /// </summary>
    static const number_format date_time2();

    /// <summary>
    /// Number format "h:mm"
    /// </summary>
    static const number_format date_time3();

    /// <summary>
    /// Number format "h:mm:ss"
    /// </summary>
    static const number_format date_time4();

    /// <summary>
    /// Number format "mm:ss"
    /// </summary>
    static const number_format date_time5();

    /// <summary>
    /// Number format "h:mm:ss"
    /// </summary>
    static const number_format date_time6();

    /// <summary>
    /// Returns true if the given format ID corresponds to a known builtin format.
    /// </summary>
    static bool is_builtin_format(std::size_t builtin_id);

    /// <summary>
    /// Returns the format with the given ID. Thows an invalid_parameter exception
    /// if builtin_id is not a valid ID.
    /// </summary>
    static const number_format &from_builtin_id(std::size_t builtin_id);

    /// <summary>
    /// Constructs a default number_format equivalent to "General"
    /// </summary>
    number_format();

    /// <summary>
    /// Constructs a number format equivalent to that returned from number_format::from_builtin_id(builtin_id).
    /// </summary>
    number_format(std::size_t builtin_id);

    /// <summary>
    /// Constructs a number format from a code string. If the string matches a builtin ID,
    /// its ID will also be set to match the builtin ID.
    /// </summary>
    number_format(const std::string &code);

    /// <summary>
    /// Constructs a number format from a code string and custom ID. Custom ID should generally
    /// be >= 164.
    /// </summary>
    number_format(const std::string &code, std::size_t custom_id);

    /// <summary>
    /// Sets the format code of this number format to format_code.
    /// </summary>
    void format_string(const std::string &format_code);

    /// <summary>
    /// Sets the format code of this number format to format_code and the ID to custom_id.
    /// </summary>
    void format_string(const std::string &format_code, std::size_t custom_id);

    /// <summary>
    /// Returns the format code this number format uses.
    /// </summary>
    std::string format_string() const;

    /// <summary>
    /// Returns true if this number format has an ID.
    /// </summary>
    bool has_id() const;

    /// <summary>
    /// Sets the ID of this number format to id.
    /// </summary>
    void id(std::size_t id);

    /// <summary>
    /// Returns the ID of this format.
    /// </summary>
    std::size_t id() const;

    /// <summary>
    /// Returns text formatted according to this number format's format code.
    /// </summary>
    std::string format(const std::string &text) const;

    /// <summary>
    /// Returns number formatted according to this number format's format code
    /// with the given base date.
    /// </summary>
    std::string format(double number, calendar base_date) const;

    /// <summary>
    /// Returns true if this format code returns a number formatted as a date.
    /// </summary>
    bool is_date_format() const;

    /// <summary>
    /// Returns true if this format is equivalent to other.
    /// </summary>
    bool operator==(const number_format &other) const;

    /// <summary>
    /// Returns true if this format is not equivalent to other.
    /// </summary>
    bool operator!=(const number_format &other) const;

private:
    /// <summary>
    /// The optional ID
    /// </summary>
    optional<std::size_t> id_;

    /// <summary>
    /// The format code
    /// </summary>
    std::string format_string_;
};

} // namespace xlnt
