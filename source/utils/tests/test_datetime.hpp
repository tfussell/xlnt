#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/xlnt.hpp>

class test_datetime : public CxxTest::TestSuite
{
public:
    void test_from_string()
    {
        xlnt::time t("10:35:45");
        TS_ASSERT_EQUALS(t.hour, 10);
        TS_ASSERT_EQUALS(t.minute, 35);
        TS_ASSERT_EQUALS(t.second, 45);
    }

    void test_to_string()
    {
        xlnt::datetime dt(2016, 7, 16, 9, 11, 32, 999999);
        TS_ASSERT_EQUALS(dt.to_string(), "2016/7/16 9:11:32:999999");
    }

    void test_carry()
    {
        // We want a time that rolls over to the next second, minute, and hour
        // Start off with a time 1 microsecond before the next hour
        xlnt::datetime dt(2016, 7, 9, 10, 59, 59, 999999);
        auto number = dt.to_number(xlnt::calendar::windows_1900);

        // Add 600 nanoseconds to the raw number which represents time as a fraction of a day
        // In other words, 6 tenths of a millionth of a sixtieth of a sixtieth of a twenty-fourth of a day
        number += (0.6 / 1000000) / 60 / 60 / 24;
        auto rollover = xlnt::datetime::from_number(number, xlnt::calendar::windows_1900);

        TS_ASSERT_EQUALS(rollover.hour, 11);
        TS_ASSERT_EQUALS(rollover.minute, 00);
        TS_ASSERT_EQUALS(rollover.second, 00);
        TS_ASSERT_EQUALS(rollover.microsecond, 00);
    }

    void test_leap_year_bug()
    {
        xlnt::datetime dt(1900, 2, 29, 0, 0, 0, 0);
        auto number = dt.to_number(xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(static_cast<int>(number), 60);
        auto converted = xlnt::datetime::from_number(number, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(dt, converted);
    }

    void test_early_date()
    {
        xlnt::date d(1900, 1, 29);
        auto number = d.to_number(xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(number, 29);
        auto converted = xlnt::date::from_number(number, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(d, converted);
    }

    void test_mac_calendar()
    {
        xlnt::date d(2016, 7, 16);
        auto number_1900 = d.to_number(xlnt::calendar::windows_1900);
        auto number_1904 = d.to_number(xlnt::calendar::mac_1904);
        TS_ASSERT_EQUALS(number_1900, 42567);
        TS_ASSERT_EQUALS(number_1904, 41105);
        auto converted_1900 = xlnt::date::from_number(number_1900, xlnt::calendar::windows_1900);
        auto converted_1904 = xlnt::date::from_number(number_1904, xlnt::calendar::mac_1904);
        TS_ASSERT_EQUALS(converted_1900, d);
        TS_ASSERT_EQUALS(converted_1904, d);
    }

    void test_operators()
    {
        xlnt::date d1(2016, 7, 16);
        xlnt::date d2(2016, 7, 16);
        xlnt::date d3(2016, 7, 15);
        TS_ASSERT_EQUALS(d1, d2);
        TS_ASSERT_DIFFERS(d1, d3);
    }
};
