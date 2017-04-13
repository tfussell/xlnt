#pragma once

#include <iostream>
#include <helpers/test_suite.hpp>

#include <xlnt/xlnt.hpp>

class test_alignment : public test_suite
{
public:
    void test_all()
    {
        xlnt::alignment alignment;

        assert(!alignment.horizontal().is_set());
        assert(!alignment.vertical().is_set());
        assert(!alignment.shrink());
        assert(!alignment.wrap());
    }
};
