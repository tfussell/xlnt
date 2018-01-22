// Copyright (c) 2014-2018 Thomas Fussell
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

/// <summary>
/// A date is a specific day specified in terms of a year, month, and day.
/// It can also be initialized as a number of days since a base date
/// using date::from_number.
/// </summary>
struct XLNT_API date
{
    /// <summary>
    /// Return the current date according to the system time.
    /// </summary>
    static date today();

    /// <summary>
    /// Return a date by adding days_since_base_year to base_date.
    /// This includes leap years.
    /// </summary>
    static date from_number(int days_since_base_year, calendar base_date);

    /// <summary>
    /// Constructs a data from a given year, month, and day.
    /// </summary>
    date(int year_, int month_, int day_);

    /// <summary>
    /// Return the number of days between this date and base_date.
    /// </summary>
    int to_number(calendar base_date) const;

    /// <summary>
    /// Calculates and returns the day of the week that this date represents in the range
    /// 0 to 6 where 0 represents Sunday.
    /// </summary>
    int weekday() const;

    /// <summary>
    /// Return true if this date is equal to comparand.
    /// </summary>
    bool operator==(const date &comparand) const;

    /// <summary>
    /// Return true if this date is equal to comparand.
    /// </summary>
    bool operator!=(const date &comparand) const;

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
};

} // namespace xlnt
