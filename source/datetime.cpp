#include <cmath>
#include <ctime>

#include "datetime.h"

namespace xlnt {

time::time(long double raw_time)
{
    double integer_part;
    double fractional_part = std::modf(raw_time, &integer_part);
    fractional_part *= 24;
    hour = (int)fractional_part;
    fractional_part = 60 * (fractional_part - hour);
    minute = (int)fractional_part;
    fractional_part = 60 * (fractional_part - minute);
    second = (int)fractional_part;
    fractional_part = 1000000 * (fractional_part - second);
    microsecond = (int)fractional_part;
}

datetime datetime::now()
{
    std::time_t raw_time = std::time(0);
    std::tm now = *std::localtime(&raw_time);
    return datetime(now.tm_year, now.tm_mon + 1, now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec);
}

} // namespace xlnt
