#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

class test_timedelta : public CxxTest::TestSuite
{
public:
    void test_from_number()
    {
        auto td = xlnt::timedelta::from_number(1.0423726852L);

        TS_ASSERT(td.days == 1);
        TS_ASSERT(td.hours == 1);
        TS_ASSERT(td.minutes == 1);
        TS_ASSERT(td.seconds == 1);
        TS_ASSERT(td.microseconds == 1);
    }

    void test_round_trip()
    {
    	long double time = 3.14159265359L;
        auto td = xlnt::timedelta::from_number(time);
        auto time_rt = td.to_number();
        TS_ASSERT_EQUALS(time, time_rt);
    }

    void test_to_number()
    {
    	xlnt::timedelta td(1, 1, 1, 1, 1);
    	TS_ASSERT_EQUALS(td.to_number(), 1.0423726852L);
    }

    void test_carry()
    {
        // We want a time that rolls over to the next second, minute, and hour
        // Start off with a time 1 microsecond before the next hour
        xlnt::timedelta td(1, 23, 59, 59, 999999);
        auto number = td.to_number();

        // Add 600 nanoseconds to the raw number which represents time as a fraction of a day
        // In other words, 6 tenths of a millionth of a sixtieth of a sixtieth of a twenty-fourth of a day
        number += (0.6 / 1000000) / 60 / 60 / 24;
        auto rollover = xlnt::timedelta::from_number(number);

        TS_ASSERT_EQUALS(rollover.days, 2);
        TS_ASSERT_EQUALS(rollover.hours, 0);
        TS_ASSERT_EQUALS(rollover.minutes, 0);
        TS_ASSERT_EQUALS(rollover.seconds, 0);
        TS_ASSERT_EQUALS(rollover.microseconds, 0);
    }
};
