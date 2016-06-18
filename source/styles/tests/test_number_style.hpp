#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include "pugixml.hpp"
#include <xlnt/xlnt.hpp>

class test_number_style : public CxxTest::TestSuite
{
public:
    void test_simple_date()
    {
        auto date = xlnt::date(2016, 6, 18);
        auto date_number = date.to_number(xlnt::calendar::windows_1900);

        xlnt::number_format nf = xlnt::number_format::date_ddmmyyyy();
        auto formatted = nf.format(date_number, xlnt::calendar::windows_1900);

        TS_ASSERT_EQUALS(formatted, "18/06/16");
    }

    void test_text_section_string()
    {
        xlnt::number_format nf;
        nf.set_format_string("General;General;General;[Magenta]\"a\"@\"b\"");

        auto formatted = nf.format("text");

        TS_ASSERT_EQUALS(formatted, "atextb");
    }
};
