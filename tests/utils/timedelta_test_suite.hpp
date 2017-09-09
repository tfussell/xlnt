// Copyright (c) 2014-2017 Thomas Fussell
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, WRISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE
//
// @license: http://www.opensource.org/licenses/mit-license.php
// @author: see AUTHORS file

#pragma once

#include <iostream>

#include <helpers/test_suite.hpp>

class timedelta_test_suite : public test_suite
{
public:
    timedelta_test_suite()
    {
        register_test(test_from_number);
        register_test(test_round_trip);
        register_test(test_to_number);
        register_test(test_carry);
    }

    void test_from_number()
    {
        auto td = xlnt::timedelta::from_number(1.0423726852L);

        xlnt_assert(td.days == 1);
        xlnt_assert(td.hours == 1);
        xlnt_assert(td.minutes == 1);
        xlnt_assert(td.seconds == 1);
        xlnt_assert(td.microseconds == 1);
    }

    void test_round_trip()
    {
        double time = 3.14159265359;
        auto td = xlnt::timedelta::from_number(time);
        auto time_rt = td.to_number();
        xlnt_assert_delta(time, time_rt, 1E-9);
    }

    void test_to_number()
    {
        xlnt::timedelta td(1, 1, 1, 1, 1);
        xlnt_assert_delta(td.to_number(), 1.0423726852L, 1E-9);
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

        xlnt_assert_equals(rollover.days, 2);
        xlnt_assert_equals(rollover.hours, 0);
        xlnt_assert_equals(rollover.minutes, 0);
        xlnt_assert_equals(rollover.seconds, 0);
        xlnt_assert_equals(rollover.microseconds, 0);
    }
};
