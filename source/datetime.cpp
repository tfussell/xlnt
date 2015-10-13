#include <cmath>
#include <ctime>

#include <xlnt/common/datetime.hpp>

namespace xlnt {

time time::from_number(long double raw_time)
{
	time result;

    double integer_part;
    double fractional_part = std::modf((double)raw_time, &integer_part);

    fractional_part *= 24;
    result.hour = (int)fractional_part;
	fractional_part = 60 * (fractional_part - result.hour);
	result.minute = (int)fractional_part;
	fractional_part = 60 * (fractional_part - result.minute);
	result.second = (int)fractional_part;
	fractional_part = 1000000 * (fractional_part - result.second);
	result.microsecond = (int)fractional_part;
	if (result.microsecond == 999999 && fractional_part - result.microsecond > 0.5)
    {
		result.microsecond = 0;
		result.second += 1;
		if (result.second == 60)
        {
			result.second = 0;
			result.minute += 1;
			if (result.minute == 60)
            {
				result.minute = 0;
				result.hour += 1;
            }
        }
    }
	return result;
}

date date::from_number(int days_since_base_year, calendar base_date)
{
    date result(0, 0, 0);

    if(base_date == calendar::mac_1904)
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

datetime datetime::from_number(long double raw_time, calendar base_date)
{
    auto date_part = date::from_number((int)raw_time, base_date);
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

long double time::to_number() const
{
    std::size_t microseconds = microsecond;
    microseconds += second * 1e6;
    microseconds += minute * 1e6 * 60;
    microseconds += hour * 1e6 * 60 * 60;
    auto number = microseconds / (24.0L * 60 * 60 * 1e6L);
    number = std::floor(number * 1e11L + 0.5L) / 1e11L;
    return number;
}

int date::to_number(calendar base_date) const
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

    if(base_date == calendar::mac_1904)
    {
        return days_since_1900 - 1462;
    }

    return days_since_1900;
}

long double datetime::to_number(calendar base_date) const
{
    return date(year, month, day).to_number(base_date)
        + time(hour, minute, second, microsecond).to_number();
}

std::string datetime::to_string(xlnt::calendar base_date) const
{
    return std::to_string(year) + "/" + std::to_string(month) + "/" + std::to_string(day) + " " +std::to_string(hour) + ":" + std::to_string(minute) + ":" + std::to_string(second) + ":" + std::to_string(microsecond);
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
    
double timedelta::to_number() const
{
    return days + hours / 24.0;
}
    
} // namespace xlnt
