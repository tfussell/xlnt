#pragma once

#include <iostream>
#include <helpers/test_suite.hpp>

class test_timedelta : public test_suite
{
public:
    void test_from_number()
    {
        auto td = xlnt::timedelta::from_number(1.0423726852L);

        assert(td.days == 1);
        assert(td.hours == 1);
        assert(td.minutes == 1);
        assert(td.seconds == 1);
        assert(td.microseconds == 1);
    }

    void test_round_trip()
    {
    	long double time = 3.14159265359L;
        auto td = xlnt::timedelta::from_number(time);
        auto time_rt = td.to_number();
        assert_equals(time, time_rt);
    }

    void test_to_number()
    {
    	xlnt::timedelta td(1, 1, 1, 1, 1);
    	assert_equals(td.to_number(), 1.0423726852L);
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

        assert_equals(rollover.days, 2);
        assert_equals(rollover.hours, 0);
        assert_equals(rollover.minutes, 0);
        assert_equals(rollover.seconds, 0);
        assert_equals(rollover.microseconds, 0);
    }
};
