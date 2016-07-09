#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/xlnt.hpp>

class test_alignment : public CxxTest::TestSuite
{
public:
    void test_all()
    {
        xlnt::alignment alignment;

        TS_ASSERT(alignment.has_horizontal());
        TS_ASSERT_EQUALS(alignment.get_horizontal(), xlnt::horizontal_alignment::general);
        alignment.set_horizontal(xlnt::horizontal_alignment::none);
        TS_ASSERT(!alignment.has_horizontal());

        TS_ASSERT(alignment.has_vertical());
        TS_ASSERT_EQUALS(alignment.get_vertical(), xlnt::vertical_alignment::bottom);
        alignment.set_vertical(xlnt::vertical_alignment::none);
        TS_ASSERT(!alignment.has_vertical());

        TS_ASSERT(!alignment.get_shrink_to_fit());
        alignment.set_shrink_to_fit(true);
        TS_ASSERT(alignment.get_shrink_to_fit());
        
        TS_ASSERT(!alignment.get_wrap_text());
        alignment.set_wrap_text(true);
        TS_ASSERT(alignment.get_wrap_text());
    }
};
