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

class index_types_test_suite : public test_suite
{
public:
    index_types_test_suite()
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
