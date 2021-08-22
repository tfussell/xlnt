// Copyright (c) 2014-2021 Thomas Fussell
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

#include <string>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/utils/calendar.hpp>

namespace xlnt {

struct date;
struct time;

/// <summary>
/// A datetime is a combination of a date and a time.
/// </summary>
struct XLNT_API datetime
{
    /// <summary>
    /// Returns the current date and time according to the system time.
    /// </summary>
    static datetime now();

    /// <summary>
    /// Returns the current date and time according to the system time.
    /// This is equivalent to datetime::now().
    /// </summary>
    static datetime today();

    /// <summary>
    /// Returns a datetime from number by converting the integer part into
    /// a date and the fractional part into a time according to date::from_number
    /// and time::from_number.
    /// </summary>
    static datetime from_number(double number, calendar base_date);

    /// <summary>
    /// Returns a datetime equivalent to the ISO-formatted string iso_string.
    /// </summary>
    static datetime from_iso_string(const std::string &iso_string);

    /// <summary>
    /// Constructs a datetime from a date and a time.
    /// </summary>
    datetime(const date &d, const time &t);

    /// <summary>
    /// Constructs a datetime from a year, month, and day plus optional hour, minute, second, and microsecond.
    /// </summary>
    datetime(int year_, int month_, int day_, int hour_ = 0, int minute_ = 0, int second_ = 0, int microsecond_ = 0);

    /// <summary>
    /// Returns a string represenation of this date and time.
    /// </summary>
    std::string to_string() const;

    /// <summary>
    /// Returns an ISO-formatted string representation of this date and time.
    /// </summary>
    std::string to_iso_string() const;

    /// <summary>
    /// Returns this datetime as a number of seconds since 1900 or 1904 (depending on base_date provided).
    /// </summary>
    double to_number(calendar base_date) const;

    /// <summary>
    /// Returns true if this datetime is equivalent to comparand.
    /// </summary>
    bool operator==(const datetime &comparand) const;

    /// <summary>
    /// Calculates and returns the day of the week that this date represents in the range
    /// 0 to 6 where 0 represents Sunday.
    /// </summary>
    int weekday() const;

    /// <summary>
    /// The year
    /// </summary>
    int year;

    /// <summary>
    /// The month
    /// </summary>
    int month;

    /// <summary>
    /// The day
    /// </summary>
    int day;

    /// <summary>
    /// The hour
    /// </summary>
    int hour;

    /// <summary>
    /// The minute
    /// </summary>
    int minute;

    /// <summary>
    /// The second
    /// </summary>
    int second;

    /// <summary>
    /// The microsecond
    /// </summary>
    int microsecond;
};

} // namespace xlnt
