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

        TS_ASSERT(!alignment.horizontal().is_set());
        TS_ASSERT(!alignment.vertical().is_set());
        TS_ASSERT(!alignment.shrink());
        TS_ASSERT(!alignment.wrap());
    }
};
