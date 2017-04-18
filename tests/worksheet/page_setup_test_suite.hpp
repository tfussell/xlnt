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

class page_setup_test_suite : public test_suite
{
public:
    page_setup_test_suite()
    {
        register_test(test_properties);
    }

    void test_properties()
    {
        xlnt::page_setup ps;

        xlnt_assert_equals(ps.paper_size(), xlnt::paper_size::letter);
        ps.paper_size(xlnt::paper_size::executive);
        xlnt_assert_equals(ps.paper_size(), xlnt::paper_size::executive);

        xlnt_assert_equals(ps.orientation(), xlnt::orientation::portrait);
        ps.orientation(xlnt::orientation::landscape);
        xlnt_assert_equals(ps.orientation(), xlnt::orientation::landscape);

        xlnt_assert(!ps.fit_to_page());
        ps.fit_to_page(true);
        xlnt_assert(ps.fit_to_page());

        xlnt_assert(!ps.fit_to_height());
        ps.fit_to_height(true);
        xlnt_assert(ps.fit_to_height());

        xlnt_assert(!ps.fit_to_width());
        ps.fit_to_width(true);
        xlnt_assert(ps.fit_to_width());

        xlnt_assert(!ps.horizontal_centered());
        ps.horizontal_centered(true);
        xlnt_assert(ps.horizontal_centered());

        xlnt_assert(!ps.vertical_centered());
        ps.vertical_centered(true);
        xlnt_assert(ps.vertical_centered());
    }
};
