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
#include <cmath>
#include <ctime>

#include <xlnt/utils/time.hpp>

namespace {

std::tm safe_localtime(std::time_t raw_time)
{
#ifdef _MSC_VER
    std::tm result;
    localtime_s(&result, &raw_time);

    return result;
#else
    return *localtime(&raw_time);
#endif
}

} // namespace

namespace xlnt {

time time::from_number(double raw_time)
{
    time result;

    double integer_part;
    double fractional_part = std::modf(static_cast<double>(raw_time), &integer_part);

    fractional_part *= 24;
    result.hour = static_cast<int>(fractional_part);
    fractional_part = 60 * (fractional_part - result.hour);
    result.minute = static_cast<int>(fractional_part);
    fractional_part = 60 * (fractional_part - result.minute);
    result.second = static_cast<int>(fractional_part);
    fractional_part = 1000000 * (fractional_part - result.second);
    result.microsecond = static_cast<int>(fractional_part);

    if (result.microsecond == 999999 && fractional_part - result.microsecond > 0.5)
    {
        result.microsecond = 0;
        result.second += 1;

        if (result.second == 60)
        {
            result.second = 0;
            result.minute += 1;

            // TODO: too much nesting
            if (result.minute == 60)
            {
                result.minute = 0;
                result.hour += 1;
            }
        }
    }

    return result;
}

time::time(int hour_, int minute_, int second_, int microsecond_)
    : hour(hour_), minute(minute_), second(second_), microsecond(microsecond_)
{
}

bool time::operator==(const time &comparand) const
{
    return hour == comparand.hour && minute == comparand.minute && second == comparand.second
        && microsecond == comparand.microsecond;
}

time::time(const std::string &time_string)
    : hour(0), minute(0), second(0), microsecond(0)
{
    std::string remaining = time_string;
    auto colon_index = remaining.find(':');
    hour = std::stoi(remaining.substr(0, colon_index));
    remaining = remaining.substr(colon_index + 1);
    colon_index = remaining.find(':');
    minute = std::stoi(remaining.substr(0, colon_index));
    colon_index = remaining.find(':');

    if (colon_index != std::string::npos)
    {
        remaining = remaining.substr(colon_index + 1);
        second = std::stoi(remaining);
    }
}

double time::to_number() const
{
    std::uint64_t microseconds = static_cast<std::uint64_t>(microsecond);
    microseconds += static_cast<std::uint64_t>(second * 1e6);
    microseconds += static_cast<std::uint64_t>(minute * 1e6 * 60);
    auto microseconds_per_hour = static_cast<std::uint64_t>(1e6) * 60 * 60;
    microseconds += static_cast<std::uint64_t>(hour) * microseconds_per_hour;
    auto number = microseconds / (24.0 * microseconds_per_hour);
    auto hundred_billion = static_cast<std::uint64_t>(1e9) * 100;
    number = std::floor(number * hundred_billion + 0.5) / hundred_billion;

    return number;
}

time time::now()
{
    std::tm now = safe_localtime(std::time(nullptr));
    return time(now.tm_hour, now.tm_min, now.tm_sec);
}

} // namespace xlnt
