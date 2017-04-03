#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/xlnt.hpp>

class test_index_types : public CxxTest::TestSuite
{
public:
    void test_bad_string_empty()
    {
        TS_ASSERT_THROWS(xlnt::column_t::column_index_from_string(""),
            xlnt::invalid_column_index);
    }
    
    void test_bad_string_too_long()
    {
        TS_ASSERT_THROWS(xlnt::column_t::column_index_from_string("ABCD"),
            xlnt::invalid_column_index);
    }
    
    void test_bad_string_numbers()
    {
        TS_ASSERT_THROWS(xlnt::column_t::column_index_from_string("123"),
            xlnt::invalid_column_index);
    }

    void test_bad_index_zero()
    {
        TS_ASSERT_THROWS(xlnt::column_t::column_string_from_index(0),
            xlnt::invalid_column_index);
    }

    void test_column_operators()
    {
        auto c1 = xlnt::column_t();
        c1 = "B";

        auto c2 = xlnt::column_t();
        const auto d = std::string("D");
        c2 = d;

        TS_ASSERT(c1 != c2);
        TS_ASSERT(c1 == static_cast<xlnt::column_t::index_t>(2));
        TS_ASSERT(c2 == d);
        TS_ASSERT(c1 != 3);
        TS_ASSERT(c1 != static_cast<xlnt::column_t::index_t>(5));
        TS_ASSERT(c1 != "D");
        TS_ASSERT(c1 != d);
        TS_ASSERT(c2 >= c1);
        TS_ASSERT(c2 > c1);
        TS_ASSERT(c1 < c2);
        TS_ASSERT(c1 <= c2);
        
        TS_ASSERT_EQUALS(--c2, 3);
        TS_ASSERT_EQUALS(c2--, 3);
        TS_ASSERT_EQUALS(c2, 2);
        
        c2 = 4;
        c1 = 3;
        
        TS_ASSERT(c2 <= 4);
        TS_ASSERT(!(c2 < 3));
        TS_ASSERT(c1 >= 3);
        TS_ASSERT(!(c1 > 4));

        TS_ASSERT(4 >= c2);
        TS_ASSERT(!(3 >= c2));
        TS_ASSERT(3 <= c1);
        TS_ASSERT(!(4 <= c1));
    }
};
