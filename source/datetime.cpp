#include <cmath>
#include <ctime>

#include "common/datetime.hpp"

namespace xlnt {

time::time(long double raw_time) : hour(0), minute(0), second(0), microsecond(0)
{
    double integer_part;
    double fractional_part = std::modf((double)raw_time, &integer_part);
    fractional_part *= 24;
    hour = (int)fractional_part;
    fractional_part = 60 * (fractional_part - hour);
    minute = (int)fractional_part;
    fractional_part = 60 * (fractional_part - minute);
    second = (int)fractional_part;
    fractional_part = 1000000 * (fractional_part - second);
    microsecond = (int)fractional_part;
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

datetime datetime::now()
{
    std::time_t raw_time = std::time(0);
    std::tm now = *std::localtime(&raw_time);
    return datetime(now.tm_year, now.tm_mon + 1, now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec);
}

} // namespace xlnt
