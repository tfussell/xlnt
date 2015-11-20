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

} // namespace xlnt
