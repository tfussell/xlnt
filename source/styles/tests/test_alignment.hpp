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

        TS_ASSERT(alignment.horizontal());
        TS_ASSERT_EQUALS(*alignment.horizontal(), xlnt::horizontal_alignment::general);
        TS_ASSERT(alignment.vertical());
        TS_ASSERT_EQUALS(*alignment.vertical(), xlnt::vertical_alignment::bottom);
        TS_ASSERT(!alignment.shrink());
        TS_ASSERT(!alignment.wrap());
    }
};
