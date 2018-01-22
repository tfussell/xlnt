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

namespace xlnt {

/// <summary>
/// A time is a specific time of the day specified in terms of an hour,
/// minute, second, and microsecond (0-999999).
/// It can also be initialized as a fraction of a day using time::from_number.
/// </summary>
struct XLNT_API time
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
    static time from_number(double number);

    /// <summary>
    /// Constructs a time object from an optional hour, minute, second, and microsecond.
    /// </summary>
    explicit time(int hour_ = 0, int minute_ = 0, int second_ = 0, int microsecond_ = 0);

    /// <summary>
    /// Constructs a time object from a string representing the time.
    /// </summary>
    explicit time(const std::string &time_string);

    /// <summary>
    /// Returns a numeric representation of the time in the range 0-1 where the value
    /// is equal to the fraction of the day elapsed.
    /// </summary>
    double to_number() const;

    /// <summary>
    /// Returns true if this time is equivalent to comparand.
    /// </summary>
    bool operator==(const time &comparand) const;

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
