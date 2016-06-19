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

        TS_ASSERT_EQUALS(ps.get_paper_size(), xlnt::paper_size::letter);
        ps.set_paper_size(xlnt::paper_size::executive);
        TS_ASSERT_EQUALS(ps.get_paper_size(), xlnt::paper_size::executive);

        TS_ASSERT_EQUALS(ps.get_orientation(), xlnt::orientation::portrait);
        ps.set_orientation(xlnt::orientation::landscape);
        TS_ASSERT_EQUALS(ps.get_orientation(), xlnt::orientation::landscape);

        TS_ASSERT(!ps.fit_to_page());
        ps.set_fit_to_page(true);
        TS_ASSERT(ps.fit_to_page());

        TS_ASSERT(!ps.fit_to_height());
        ps.set_fit_to_height(true);
        TS_ASSERT(ps.fit_to_height());

        TS_ASSERT(!ps.fit_to_width());
        ps.set_fit_to_width(true);
        TS_ASSERT(ps.fit_to_width());

        TS_ASSERT(!ps.get_horizontal_centered());
        ps.set_horizontal_centered(true);
        TS_ASSERT(ps.get_horizontal_centered());

        TS_ASSERT(!ps.get_vertical_centered());
        ps.set_vertical_centered(true);
        TS_ASSERT(ps.get_vertical_centered());
    }
};
