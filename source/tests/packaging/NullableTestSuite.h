#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include "../../misc/nullable.h"

class NullableTestSuite : public CxxTest::TestSuite
{
public:
    NullableTestSuite() : integer_non_null_5(5), double_non_null_5(5.0), double_non_null_6(6.0)
    {

    }

    void test_has_value()
    {
        TS_ASSERT(!integer_null.has_value())
        TS_ASSERT(integer_non_null_5.has_value())
    }

    void test_get_value()
    {
        TS_ASSERT_THROWS_ANYTHING(integer_null.get_value())
        TS_ASSERT_THROWS_NOTHING(integer_non_null_5.get_value())
        TS_ASSERT_EQUALS(integer_non_null_5.get_value(), 5)
    }

    void test_equality()
    {
        TS_ASSERT_EQUALS(integer_null, integer_null_2)
        TS_ASSERT_DIFFERS(integer_null, integer_non_null_5)
        TS_ASSERT_DIFFERS(double_non_null_5, double_non_null_6)
    }

    void test_assignment()
    {
        xlnt::nullable<double> double_non_null_6_assigned;
        double_non_null_6_assigned = double_non_null_6;
        TS_ASSERT(double_non_null_6_assigned.has_value())
        TS_ASSERT_EQUALS(double_non_null_6_assigned.get_value(), 6.0)
    }

    void test_copy_constructor()
    {
        xlnt::nullable<int> integer_non_null_5_copy(integer_non_null_5);
        TS_ASSERT(integer_non_null_5_copy.has_value())
        TS_ASSERT_EQUALS(integer_non_null_5_copy.get_value(), 5)
    }

private:
    xlnt::nullable<int> integer_null;
    xlnt::nullable<int> integer_null_2;
    xlnt::nullable<int> integer_non_null_5;
    xlnt::nullable<double> double_null;
    xlnt::nullable<double> double_non_null_5;
    xlnt::nullable<double> double_non_null_6;
};
