#include <cmath>
#include <ctime>

#include <xlnt/utils/timedelta.hpp>

namespace xlnt {

timedelta::timedelta() : timedelta(0, 0, 0, 0, 0)
{
}

timedelta::timedelta(int days_, int hours_, int minutes_, int seconds_, int microseconds_)
        : days(days_), 
          hours(hours_), 
          minutes(minutes_), 
          seconds(seconds_), 
          microseconds(microseconds_)
    {
    }

double timedelta::to_number() const
{
    std::uint64_t total_microseconds = static_cast<std::uint64_t>(microseconds);
    total_microseconds += static_cast<std::uint64_t>(seconds * 1e6);
    total_microseconds += static_cast<std::uint64_t>(minutes * 1e6 * 60);
    auto microseconds_per_hour = static_cast<std::uint64_t>(1e6) * 60 * 60;
    total_microseconds += static_cast<std::uint64_t>(hours) * microseconds_per_hour;
    auto number = total_microseconds / (24.0L * microseconds_per_hour);
    auto hundred_billion = static_cast<std::uint64_t>(1e9) * 100;
    number = std::floor(number * hundred_billion + 0.5L) / hundred_billion;
    number += days;
    
    return number;
}

timedelta timedelta::from_number(long double raw_time)
{
    timedelta result;

    double integer_part;
    double fractional_part = std::modf((double)raw_time, &integer_part);

    result.days = integer_part;

    fractional_part *= 24;
    result.hours = (int)fractional_part;
    fractional_part = 60 * (fractional_part - result.hours);
    result.minutes = (int)fractional_part;
    fractional_part = 60 * (fractional_part - result.minutes);
    result.seconds = (int)fractional_part;
    fractional_part = 1000000 * (fractional_part - result.seconds);
    result.microseconds = (int)fractional_part;
    
    if (result.microseconds == 999999 && fractional_part - result.microseconds > 0.5)
    {
        result.microseconds = 0;
        result.seconds += 1;
        
        if (result.seconds == 60)
        {
            result.seconds = 0;
            result.minutes += 1;
            
            //TODO: too much nesting
            if (result.minutes == 60)
            {
                result.minutes = 0;
                result.hours += 1;
            }
        }
    }
    
    return result;
}

} // namespace xlnt
