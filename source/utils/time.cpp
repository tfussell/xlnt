#include <cmath>
#include <ctime>

#include <xlnt/utils/time.hpp>

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
            
            //TODO: too much nesting
            if (result.minute == 60)
            {
                result.minute = 0;
                result.hour += 1;
            }
        }
    }
    
    return result;
}

bool time::operator==(const time &comparand) const
{
    return hour == comparand.hour && minute == comparand.minute && second == comparand.second &&
           microsecond == comparand.microsecond;
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
    
    if (colon_index != std::string::npos)
    {
        remaining = remaining.substr(colon_index + 1);
        second = std::stoi(remaining);
    }
}

long double time::to_number() const
{
    std::uint64_t microseconds = static_cast<std::uint64_t>(microsecond);
    microseconds += static_cast<std::uint64_t>(second * 1e6);
    microseconds += static_cast<std::uint64_t>(minute * 1e6 * 60);
	auto microseconds_per_hour = static_cast<std::uint64_t>(1e6) * 60 * 60;
    microseconds += static_cast<std::uint64_t>(hour) * microseconds_per_hour;
    auto number = microseconds / (24.0L * microseconds_per_hour);
	auto hundred_billion = static_cast<std::uint64_t>(1e9) * 100;
    number = std::floor(number * hundred_billion + 0.5L) / hundred_billion;
    
    return number;
}

} // namespace xlnt
