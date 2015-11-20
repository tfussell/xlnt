#include <cmath>
#include <ctime>

#include <xlnt/utils/timedelta.hpp>

namespace xlnt {

double timedelta::to_number() const
{
    return days + hours / 24.0;
}

timedelta timedelta::from_number(long double number)
{
    int days = static_cast<int>(number);
    number -= days;
    number *= 24;
    int hours = static_cast<int>(number);
    number -= hours;
    number *= 60;
    int minutes = static_cast<int>(number);
    number -= minutes;
    number *= 60;
    int seconds = static_cast<int>(number);
    number -= seconds;
    number *= 1000000;
    int microseconds = static_cast<int>(number + 0.5);
    
    return timedelta(days, hours, minutes, seconds, microseconds);
}

} // namespace xlnt
