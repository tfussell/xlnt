#pragma once

#include <cxxtest/TestSuite.h>

class test_timedelta : public CxxTest::TestSuite
{
public:
    void test_from_number()
    {
        auto td = xlnt::timedelta::from_number(1.0);
        TS_ASSERT(td.days == 1);
        TS_ASSERT(td.hours == 0);
        TS_ASSERT(td.minutes == 0);
        TS_ASSERT(td.seconds == 0);
        TS_ASSERT(td.microseconds == 0);
    }
};
