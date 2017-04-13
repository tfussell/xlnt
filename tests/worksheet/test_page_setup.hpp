#pragma once

#include <iostream>
#include <helpers/test_suite.hpp>

#include <xlnt/xlnt.hpp>

class test_page_setup : public test_suite
{
public:
    void test_properties()
    {
        xlnt::page_setup ps;

        assert_equals(ps.paper_size(), xlnt::paper_size::letter);
        ps.paper_size(xlnt::paper_size::executive);
        assert_equals(ps.paper_size(), xlnt::paper_size::executive);

        assert_equals(ps.orientation(), xlnt::orientation::portrait);
        ps.orientation(xlnt::orientation::landscape);
        assert_equals(ps.orientation(), xlnt::orientation::landscape);

        assert(!ps.fit_to_page());
        ps.fit_to_page(true);
        assert(ps.fit_to_page());

        assert(!ps.fit_to_height());
        ps.fit_to_height(true);
        assert(ps.fit_to_height());

        assert(!ps.fit_to_width());
        ps.fit_to_width(true);
        assert(ps.fit_to_width());

        assert(!ps.horizontal_centered());
        ps.horizontal_centered(true);
        assert(ps.horizontal_centered());

        assert(!ps.vertical_centered());
        ps.vertical_centered(true);
        assert(ps.vertical_centered());
    }
};
