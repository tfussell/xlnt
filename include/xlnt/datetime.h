#pragma once

namespace xlnt {

struct date
{
    date(int year, int month, int day)
        : year(year), month(month), day(day)
    {
    }

    int year;
    int month;
    int day;
};

struct time
{
    time(int hour = 0, int minute = 0, int second = 0, int microsecond = 0)
        : hour(hour), minute(minute), second(second), microsecond(microsecond)
    {
    }

    explicit time(long double number);

    int hour;
    int minute;
    int second;
    int microsecond;
};

struct datetime
{
    static datetime now();

    datetime(int year, int month, int day, int hour = 0, int minute = 0, int second = 0, int microsecond = 0)
        : year(year), month(month), day(day), hour(hour), minute(minute), second(second), microsecond(microsecond)
    {
    }

    int year;
    int month;
    int day;
    int hour;
    int minute;
    int second;
    int microsecond;
};

} // namespace xlnt
