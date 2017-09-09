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

namespace xlnt {

/// <summary>
/// Represents a span of time between two datetimes. This is
/// not fully supported yet throughout the library.
/// </summary>
struct XLNT_API timedelta
{
    /// <summary>
    /// Returns a timedelta from a number representing the factional number of days elapsed.
    /// </summary>
    static timedelta from_number(double number);

    /// <summary>
    /// Constructs a timedelta equal to zero.
    /// </summary>
    timedelta();

    /// <summary>
    /// Constructs a timedelta from a number of days, hours, minutes, seconds, and microseconds.
    /// </summary>
    timedelta(int days_, int hours_, int minutes_, int seconds_, int microseconds_);

    /// <summary>
    /// Returns a numeric representation of this timedelta as a fractional number of days.
    /// </summary>
    double to_number() const;

    /// <summary>
    /// The days
    /// </summary>
    int days;

    /// <summary>
    /// The hours
    /// </summary>
    int hours;

    /// <summary>
    /// The minutes
    /// </summary>
    int minutes;

    /// <summary>
    /// The seconds
    /// </summary>
    int seconds;

    /// <summary>
    /// The microseconds
    /// </summary>
    int microseconds;
};

} // namespace xlnt
