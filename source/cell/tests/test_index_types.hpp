#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/xlnt.hpp>

class test_index_types : public CxxTest::TestSuite
{
public:
    void test_bad_string()
    {
        TS_ASSERT_THROWS(xlnt::column_t::column_index_from_string(""), xlnt::invalid_column_string_index);
        TS_ASSERT_THROWS(xlnt::column_t::column_index_from_string("ABCD"), xlnt::invalid_column_string_index);
        TS_ASSERT_THROWS(xlnt::column_t::column_index_from_string("123"), xlnt::invalid_column_string_index);
    }
    
    void test_bad_column()
    {
        TS_ASSERT_THROWS(xlnt::column_t::column_string_from_index(0), xlnt::invalid_column_string_index);
    }

    void test_column_operators()
    {
        xlnt::column_t c1;
        c1 = "B";
        xlnt::column_t c2;
        std::string d = "D";
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
        
        TS_ASSERT_EQUALS(c1 * c2, 8);
        TS_ASSERT_EQUALS(c2 / c1, 2);
        TS_ASSERT_EQUALS(c2 % c1, 0);
        
        c1 *= c2;
        TS_ASSERT_EQUALS(c1, 8);
        c1 /= c2;
        TS_ASSERT_EQUALS(c1, 2);
        c1 = 3;
        c1 %= c2;
        TS_ASSERT_EQUALS(c1, 3);
        
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
