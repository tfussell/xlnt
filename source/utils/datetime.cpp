// Copyright (c) 2014-2016 Thomas Fussell
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

#include <xlnt/utils/datetime.hpp>
#include <xlnt/utils/date.hpp>
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

datetime datetime::from_number(long double raw_time, calendar base_date)
{
    auto date_part = date::from_number((int)raw_time, base_date);
    auto time_part = time::from_number(raw_time);
    
    return datetime(date_part.year, date_part.month, date_part.day, time_part.hour, time_part.minute, time_part.second,
                    time_part.microsecond);
}

bool datetime::operator==(const datetime &comparand) const
{
    return year == comparand.year && month == comparand.month && day == comparand.day && hour == comparand.hour &&
           minute == comparand.minute && second == comparand.second && microsecond == comparand.microsecond;
}

long double datetime::to_number(calendar base_date) const
{
    return date(year, month, day).to_number(base_date) + time(hour, minute, second, microsecond).to_number();
}

std::string datetime::to_string(xlnt::calendar /*base_date*/) const
{
    return std::to_string(year) + "/" + std::to_string(month) + "/" + std::to_string(day) + " " + std::to_string(hour) +
           ":" + std::to_string(minute) + ":" + std::to_string(second) + ":" + std::to_string(microsecond);
}

datetime datetime::now()
{
	std::tm now = safe_localtime(std::time(0));

    return datetime(1900 + now.tm_year, now.tm_mon + 1, now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec);
}

datetime datetime::today()
{
    return now();
}

datetime::datetime(int year_, int month_, int day_, int hour_, int minute_, int second_, int microsecond_)
    : year(year_),
      month(month_),
      day(day_),
      hour(hour_),
      minute(minute_),
      second(second_),
      microsecond(microsecond_)
{
}

} // namespace xlnt
