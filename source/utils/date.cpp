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
#include <cmath>
#include <ctime>

#include <xlnt/utils/date.hpp>

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

date::date(int year_, int month_, int day_)
    : year(year_), month(month_), day(day_)
{
}

date date::from_number(int days_since_base_year, calendar base_date)
{
    date result(0, 0, 0);

    if (base_date == calendar::mac_1904)
    {
        days_since_base_year += 1462;
    }

    if (days_since_base_year == 60)
    {
        result.day = 29;
        result.month = 2;
        result.year = 1900;

        return result;
    }
    else if (days_since_base_year < 60)
    {
        days_since_base_year++;
    }

    int l = days_since_base_year + 68569 + 2415019;
    int n = int((4 * l) / 146097);
    l = l - int((146097 * n + 3) / 4);
    int i = int((4000 * (l + 1)) / 1461001);
    l = l - int((1461 * i) / 4) + 31;
    int j = int((80 * l) / 2447);
    result.day = l - int((2447 * j) / 80);
    l = int(j / 11);
    result.month = j + 2 - (12 * l);
    result.year = 100 * (n - 49) + i + l;

    return result;
}

bool date::operator==(const date &comparand) const
{
    return year == comparand.year && month == comparand.month && day == comparand.day;
}

bool date::operator!=(const date &comparand) const
{
    return !(*this == comparand);
}

int date::to_number(calendar base_date) const
{
    if (day == 29 && month == 2 && year == 1900)
    {
        return 60;
    }

    int days_since_1900 = int((1461 * (year + 4800 + int((month - 14) / 12))) / 4)
        + int((367 * (month - 2 - 12 * ((month - 14) / 12))) / 12)
        - int((3 * (int((year + 4900 + int((month - 14) / 12)) / 100))) / 4) + day - 2415019 - 32075;

    if (days_since_1900 <= 60)
    {
        days_since_1900--;
    }

    if (base_date == calendar::mac_1904)
    {
        return days_since_1900 - 1462;
    }

    return days_since_1900;
}

date date::today()
{
    std::tm now = safe_localtime(std::time(nullptr));
    return date(1900 + now.tm_year, now.tm_mon + 1, now.tm_mday);
}

int date::weekday() const
{
    std::tm tm = std::tm();
    tm.tm_mday = day;
    tm.tm_mon = month - 1;
    tm.tm_year = year - 1900;
    std::time_t time = std::mktime(&tm);

    return safe_localtime(time).tm_wday;
}

} // namespace xlnt
