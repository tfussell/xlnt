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
#include <xlnt/xlnt.hpp>

class datetime_test_suite : public test_suite
{
public:
    datetime_test_suite()
    {
        register_test(test_from_string);
        register_test(test_to_string);
        register_test(test_carry);
        register_test(test_leap_year_bug);
        register_test(test_early_date);
        register_test(test_mac_calendar);
        register_test(test_operators);
    }

    void test_from_string()
    {
        xlnt::time t("10:35:45");
        xlnt_assert_equals(t.hour, 10);
        xlnt_assert_equals(t.minute, 35);
        xlnt_assert_equals(t.second, 45);
    }

    void test_to_string()
    {
        xlnt::datetime dt(2016, 7, 16, 9, 11, 32, 999999);
        xlnt_assert_equals(dt.to_string(), "2016/7/16 9:11:32:999999");
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

        xlnt_assert_equals(rollover.hour, 11);
        xlnt_assert_equals(rollover.minute, 00);
        xlnt_assert_equals(rollover.second, 00);
        xlnt_assert_equals(rollover.microsecond, 00);
    }

    void test_leap_year_bug()
    {
        xlnt::datetime dt(1900, 2, 29, 0, 0, 0, 0);
        auto number = dt.to_number(xlnt::calendar::windows_1900);
        xlnt_assert_equals(static_cast<int>(number), 60);
        auto converted = xlnt::datetime::from_number(number, xlnt::calendar::windows_1900);
        xlnt_assert_equals(dt, converted);
    }

    void test_early_date()
    {
        xlnt::date d(1900, 1, 29);
        auto number = d.to_number(xlnt::calendar::windows_1900);
        xlnt_assert_equals(number, 29);
        auto converted = xlnt::date::from_number(number, xlnt::calendar::windows_1900);
        xlnt_assert_equals(d, converted);
    }

    void test_mac_calendar()
    {
        xlnt::date d(2016, 7, 16);
        auto number_1900 = d.to_number(xlnt::calendar::windows_1900);
        auto number_1904 = d.to_number(xlnt::calendar::mac_1904);
        xlnt_assert_equals(number_1900, 42567);
        xlnt_assert_equals(number_1904, 41105);
        auto converted_1900 = xlnt::date::from_number(number_1900, xlnt::calendar::windows_1900);
        auto converted_1904 = xlnt::date::from_number(number_1904, xlnt::calendar::mac_1904);
        xlnt_assert_equals(converted_1900, d);
        xlnt_assert_equals(converted_1904, d);
    }

    void test_operators()
    {
        xlnt::date d1(2016, 7, 16);
        xlnt::date d2(2016, 7, 16);
        xlnt::date d3(2016, 7, 15);
        xlnt_assert_equals(d1, d2);
        xlnt_assert_differs(d1, d3);
    }
};
