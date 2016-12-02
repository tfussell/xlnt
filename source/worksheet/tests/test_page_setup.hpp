#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/xlnt.hpp>

class test_page_setup : public CxxTest::TestSuite
{
public:
    void test_properties()
    {
        xlnt::page_setup ps;

        TS_ASSERT_EQUALS(ps.paper_size(), xlnt::paper_size::letter);
        ps.paper_size(xlnt::paper_size::executive);
        TS_ASSERT_EQUALS(ps.paper_size(), xlnt::paper_size::executive);

        TS_ASSERT_EQUALS(ps.orientation(), xlnt::orientation::portrait);
        ps.orientation(xlnt::orientation::landscape);
        TS_ASSERT_EQUALS(ps.orientation(), xlnt::orientation::landscape);

        TS_ASSERT(!ps.fit_to_page());
        ps.fit_to_page(true);
        TS_ASSERT(ps.fit_to_page());

        TS_ASSERT(!ps.fit_to_height());
        ps.fit_to_height(true);
        TS_ASSERT(ps.fit_to_height());

        TS_ASSERT(!ps.fit_to_width());
        ps.fit_to_width(true);
        TS_ASSERT(ps.fit_to_width());

        TS_ASSERT(!ps.horizontal_centered());
        ps.horizontal_centered(true);
        TS_ASSERT(ps.horizontal_centered());

        TS_ASSERT(!ps.vertical_centered());
        ps.vertical_centered(true);
        TS_ASSERT(ps.vertical_centered());
    }
};
