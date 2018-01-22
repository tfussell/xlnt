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
#include <cmath>
#include <ctime>

#include <xlnt/utils/timedelta.hpp>

namespace xlnt {

timedelta::timedelta()
    : timedelta(0, 0, 0, 0, 0)
{
}

timedelta::timedelta(int days_, int hours_, int minutes_, int seconds_, int microseconds_)
    : days(days_), hours(hours_), minutes(minutes_), seconds(seconds_), microseconds(microseconds_)
{
}

double timedelta::to_number() const
{
    std::uint64_t total_microseconds = static_cast<std::uint64_t>(microseconds);
    total_microseconds += static_cast<std::uint64_t>(seconds * 1e6);
    total_microseconds += static_cast<std::uint64_t>(minutes * 1e6 * 60);
    auto microseconds_per_hour = static_cast<std::uint64_t>(1e6) * 60 * 60;
    total_microseconds += static_cast<std::uint64_t>(hours) * microseconds_per_hour;
    auto number = total_microseconds / (24.0 * microseconds_per_hour);
    auto hundred_billion = static_cast<std::uint64_t>(1e9) * 100;
    number = std::floor(number * hundred_billion + 0.5) / hundred_billion;
    number += days;

    return number;
}

timedelta timedelta::from_number(double raw_time)
{
    timedelta result;

    result.days = static_cast<int>(raw_time);
    double fractional_part = raw_time - result.days;

    fractional_part *= 24;
    result.hours = static_cast<int>(fractional_part);
    fractional_part = 60 * (fractional_part - result.hours);
    result.minutes = static_cast<int>(fractional_part);
    fractional_part = 60 * (fractional_part - result.minutes);
    result.seconds = static_cast<int>(fractional_part);
    fractional_part = 1000000 * (fractional_part - result.seconds);
    result.microseconds = static_cast<int>(fractional_part);

    if (result.microseconds == 999999 && fractional_part - result.microseconds > 0.5)
    {
        result.microseconds = 0;
        result.seconds += 1;

        if (result.seconds == 60)
        {
            result.seconds = 0;
            result.minutes += 1;

            // TODO: too much nesting
            if (result.minutes == 60)
            {
                result.minutes = 0;
                result.hours += 1;

                if (result.hours == 24)
                {
                    result.hours = 0;
                    result.days += 1;
                }
            }
        }
    }

    return result;
}

} // namespace xlnt
