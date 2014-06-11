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
    return time(hour, minute, second, microsecond);
}

date date::from_number(long double number)
{
    int year = (int)number / 365;
    number -= year * 365;
    int month = (int)number / 30;
    number -= month * 30;
    int day = (int)number;
    return date(year, month, day + 1);
}

datetime datetime::from_number(long double raw_time)
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
    int year = (int)integer_part / 365;
    integer_part -= year * 365;
    int month = (int)integer_part / 30;
    integer_part -= month * 30;
    int day = (int)integer_part;
    return datetime(year + 1900, month, day + 1, hour, minute, second, microsecond);
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

double date::to_number() const
{
    double number = day + month * 30 + year * 365;
    return number;
}

double datetime::to_number() const
{
    double number = microsecond;
    number /= 1000000;
    number += second;
    number /= 60;
    number += minute;
    number /= 60;
    number += hour;
    number /= 24;
    number += (day - 1) + month * 30 + (year - 1900) * 365;
    return number;
}

date date::today()
{
    std::time_t raw_time = std::time(0);
    std::tm now = *std::localtime(&raw_time);
    return date(now.tm_year, now.tm_mon + 1, now.tm_mday);
}

datetime datetime::now()
{
    std::time_t raw_time = std::time(0);
    std::tm now = *std::localtime(&raw_time);
    return datetime(now.tm_year, now.tm_mon + 1, now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec);
}

} // namespace xlnt
