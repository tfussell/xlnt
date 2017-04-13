#pragma once

#include <iostream>

#include <helpers/test_suite.hpp>
#include <xlnt/xlnt.hpp>

class test_index_types : public test_suite
{
public:
    test_index_types()
    {
	register_test(test_bad_string_empty);
	register_test(test_bad_string_too_long);
	register_test(test_bad_string_numbers);
	register_test(test_bad_index_zero);
	register_test(test_column_operators);
    }

    void test_bad_string_empty()
    {
        assert_throws(xlnt::column_t::column_index_from_string(""),
            xlnt::invalid_column_index);
    }
    
    void test_bad_string_too_long()
    {
        assert_throws(xlnt::column_t::column_index_from_string("ABCD"),
            xlnt::invalid_column_index);
    }
    
    void test_bad_string_numbers()
    {
        assert_throws(xlnt::column_t::column_index_from_string("123"),
            xlnt::invalid_column_index);
    }

    void test_bad_index_zero()
    {
        assert_throws(xlnt::column_t::column_string_from_index(0),
            xlnt::invalid_column_index);
    }

    void test_column_operators()
    {
        auto c1 = xlnt::column_t();
        c1 = "B";

        auto c2 = xlnt::column_t();
        const auto d = std::string("D");
        c2 = d;

        assert(c1 != c2);
        assert(c1 == static_cast<xlnt::column_t::index_t>(2));
        assert(c2 == d);
        assert(c1 != 3);
        assert(c1 != static_cast<xlnt::column_t::index_t>(5));
        assert(c1 != "D");
        assert(c1 != d);
        assert(c2 >= c1);
        assert(c2 > c1);
        assert(c1 < c2);
        assert(c1 <= c2);
        
        assert_equals(--c2, 3);
        assert_equals(c2--, 3);
        assert_equals(c2, 2);
        
        c2 = 4;
        c1 = 3;
        
        assert(c2 <= 4);
        assert(!(c2 < 3));
        assert(c1 >= 3);
        assert(!(c1 > 4));

        assert(4 >= c2);
        assert(!(3 >= c2));
        assert(3 <= c1);
        assert(!(4 <= c1));
    }
};
