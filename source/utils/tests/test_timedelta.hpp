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
};
