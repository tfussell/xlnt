#include <cmath>
#include <ctime>

#include "common/datetime.hpp"

namespace xlnt {

time time::from_number(long double raw_time)
{
    double integer_part;
    double fractional_part = std::modf((double)raw_time, &integer_part);
    fractional_part *= 24;
    int hour = (int)fractional_part;
    fractional_part = 60 * (fractional_part - hour);
    int minute = (int)fractional_part;
    fractional_part = 60 * (fractional_part - minute);
    int second = (int)fractional_part;
    fractional_part = 1000000 * (fractional_part - second);
    int microsecond = (int)fractional_part;
    if(microsecond == 999999 && fractional_part - microsecond > 0.5)
    {
        microsecond = 0;
        second += 1;
        if(second == 60)
        {
            second = 0;
            minute += 1;
            if(minute == 60)
            {
                minute = 0;
                hour += 1;
            }
        }
    }
    return time(hour, minute, second, microsecond);
}

date date::from_number(int days_since_base_year, int base_year)
{
    date result(0, 0, 0);

    if(base_year == 1904)
    {
        days_since_base_year += 1462;
    }

    if(days_since_base_year == 60)
    {
        result.day = 29;
        result.month = 2;
        result.year = 1900;
    }
    else if(days_since_base_year < 60)
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

datetime datetime::from_number(long double raw_time, int base_year)
{
    auto date_part = date::from_number((int)raw_time, base_year);
    auto time_part = time::from_number(raw_time);
    return datetime(date_part.year, date_part.month, date_part.day, time_part.hour, time_part.minute, time_part.second, time_part.microsecond);
}

bool date::operator==(const date &comparand) const
{
    return year == comparand.year
        && month == comparand.month
        && day == comparand.day;
}

bool time::operator==(const time &comparand) const
{
    return hour == comparand.hour
        && minute == comparand.minute
        && second == comparand.second
        && microsecond == comparand.microsecond;
}

bool datetime::operator==(const datetime &comparand) const
{
    return year == comparand.year
        && month == comparand.month
        && day == comparand.day
        && hour == comparand.hour
        && minute == comparand.minute
        && second == comparand.second
        && microsecond == comparand.microsecond;
}

time::time(const std::string &time_string) : hour(0), minute(0), second(0), microsecond(0)
{
    std::string remaining = time_string;
    auto colon_index = remaining.find(':');
    hour = std::stoi(remaining.substr(0, colon_index));
    remaining = remaining.substr(colon_index + 1);
    colon_index = remaining.find(':');
    minute = std::stoi(remaining.substr(0, colon_index));
    colon_index = remaining.find(':');
    if(colon_index != std::string::npos)
    {
	remaining = remaining.substr(colon_index + 1);    
	second = std::stoi(remaining);
    }
}

double time::to_number() const
{
    double number = microsecond;
    number /= 1000000;
    number += second;
    number /= 60;
    number += minute;
    number /= 60;
    number += hour;
    number /= 24;
    return number;
}

int date::to_number(int base_year) const
{
    if(day == 29 && month == 2 && year == 1900)
    {
        return 60;
    }

    int days_since_1900 = int((1461 * (year + 4800 + int((month - 14) / 12))) / 4)
        + int((367 * (month - 2 - 12 * ((month - 14) / 12))) / 12)
        - int((3 * (int((year + 4900 + int((month - 14) / 12)) / 100))) / 4)
        + day - 2415019 - 32075;

    if(days_since_1900 <= 60)
    {
        days_since_1900--;
    }

    if(base_year == 1904)
    {
        return days_since_1900 - 1461;
    }

    return days_since_1900;
}

double datetime::to_number(int base_year) const
{
    return date(year, month, day).to_number(base_year) 
        + time(hour, minute, second, microsecond).to_number();
}

date date::today()
{
    std::time_t raw_time = std::time(0);
    std::tm now = *std::localtime(&raw_time);
    return date(1900 + now.tm_year, now.tm_mon + 1, now.tm_mday);
}

datetime datetime::now()
{
    std::time_t raw_time = std::time(0);
    std::tm now = *std::localtime(&raw_time);
    return datetime(1900 + now.tm_year, now.tm_mon + 1, now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec);
}

} // namespace xlnt
