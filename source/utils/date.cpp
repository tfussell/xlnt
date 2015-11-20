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

int date::to_number(calendar base_date) const
{
    if (day == 29 && month == 2 && year == 1900)
    {
        return 60;
    }

    int days_since_1900 = int((1461 * (year + 4800 + int((month - 14) / 12))) / 4) +
                          int((367 * (month - 2 - 12 * ((month - 14) / 12))) / 12) -
                          int((3 * (int((year + 4900 + int((month - 14) / 12)) / 100))) / 4) + day - 2415019 - 32075;

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
    std::tm now = safe_localtime(std::time(0));

    return date(1900 + now.tm_year, now.tm_mon + 1, now.tm_mday);
}

} // namespace xlnt
