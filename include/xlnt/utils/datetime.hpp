// Copyright (c) 2014-2017 Thomas Fussell
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
    /// Return the current date and time according to the system time.
    /// </summary>
    static datetime now();

    /// <summary>
    /// Return the current date and time according to the system time.
    /// This is equivalent to datetime::now().
    /// </summary>
    static datetime today();

    /// <summary>
    /// Return a datetime from number by converting the integer part into
    /// a date and the fractional part into a time according to date::from_number
    /// and time::from_number.
    /// </summary>
    static datetime from_number(long double number, calendar base_date);

    /// <summary>
    ///
    /// </summary>
    static datetime from_iso_string(const std::string &iso_string);

    /// <summary>
    ///
    /// </summary>
    datetime(const date &d, const time &t);

    /// <summary>
    ///
    /// </summary>
    datetime(int year_, int month_, int day_, int hour_ = 0, int minute_ = 0, int second_ = 0, int microsecond_ = 0);

    /// <summary>
    ///
    /// </summary>
    std::string to_string() const;

    /// <summary>
    ///
    /// </summary>
    std::string to_iso_string() const;

    /// <summary>
    ///
    /// </summary>
    long double to_number(calendar base_date) const;

    /// <summary>
    ///
    /// </summary>
    bool operator==(const datetime &comparand) const;

    /// <summary>
    ///
    /// </summary>
    int weekday() const;

    /// <summary>
    ///
    /// </summary>
    int year;

    /// <summary>
    ///
    /// </summary>
    int month;

    /// <summary>
    ///
    /// </summary>
    int day;

    /// <summary>
    ///
    /// </summary>
    int hour;

    /// <summary>
    ///
    /// </summary>
    int minute;

    /// <summary>
    ///
    /// </summary>
    int second;

    /// <summary>
    ///
    /// </summary>
    int microsecond;
};

} // namespace xlnt
