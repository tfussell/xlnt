// Copyright (c) 2015 Thomas Fussell
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

namespace xlnt {

/// <summary>
/// An enumeration of possible base dates.
/// Dates in Excel are stored as days since this base date.
/// </summary>
enum class calendar
{
    windows_1900,
    mac_1904
};

/// <summary>
/// A date is a specific day specified in terms of a year, month, and day.
/// It can also be initialized as a number of days since a base date
/// using date::from_number.
/// </summary>
struct date
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

    date(int year_, int month_, int day_)
        : year(year_), month(month_), day(day_)
    {
    }

    /// <summary>
    /// Return the number of days between this date and base_date.
    /// </summary>
    int to_number(calendar base_date) const;
    
    /// <summary>
    /// Return true if this date is equal to comparand.
    /// </summary>
    bool operator==(const date &comparand) const;

    int year;
    int month;
    int day;
};

/// <summary>
/// A time is a specific time of the day specified in terms of an hour,
/// minute, second, and microsecond (0-999999).
/// It can also be initialized as a fraction of a day using time::from_number.
/// </summary>
struct time
{
    /// <summary>
    /// Return the current time according to the system time.
    /// </summary>
    static time now();
    
    /// <summary>
    /// Return a time from a number representing a fraction of a day.
    /// The integer part of number will be ignored.
    /// 0.5 would return time(12, 0, 0, 0) or noon, halfway through the day.
    /// </summary>
    static time from_number(long double number);

    explicit time(int hour_ = 0, int minute_ = 0, int second_ = 0, int microsecond_ = 0)
        : hour(hour_), minute(minute_), second(second_), microsecond(microsecond_)
    {
    }
    explicit time(const std::string &time_string);

    long double to_number() const;
    bool operator==(const time &comparand) const;

    int hour;
    int minute;
    int second;
    int microsecond;
};

struct datetime
{
    /// <summary>
    /// Return the current date and time according to the system time.
    /// </summary>
    static datetime now();
    
    /// <summary>
    /// Return the current date and time according to the system time.
    /// This is equivalent to datetime::now().
    /// </summary>
    static datetime today() { return now(); }
    
    /// <summary>
    /// Return a datetime from number by converting the integer part into
    /// a date and the fractional part into a time according to date::from_number
    /// and time::from_number.
    /// </summary>
    static datetime from_number(long double number, calendar base_date);

    datetime(int year_, int month_, int day_, int hour_ = 0, int minute_ = 0, int second_ = 0, int microsecond_ = 0)
        : year(year_), month(month_), day(day_), hour(hour_), minute(minute_), second(second_), microsecond(microsecond_)
    {
    }

    std::string to_string(calendar base_date) const;
    long double to_number(calendar base_date) const;
    bool operator==(const datetime &comparand) const;

    int year;
    int month;
    int day;
    int hour;
    int minute;
    int second;
    int microsecond;
};
    
/// <summary>
/// Represents a span of time between two datetimes. This is
/// not fully supported yet.
/// </summary>
struct timedelta
{
    timedelta(int days_, int hours_, int minutes_ = 0, int seconds_ = 0, int microseconds_ = 0)
    : days(days_),
      hours(hours_),
      minutes(minutes_),
      seconds(seconds_),
      microseconds(microseconds_)
    {
    }
    
    double to_number() const;
    
    int days;
    int hours;
    int minutes;
    int seconds;
    int microseconds;
};

} // namespace xlnt
