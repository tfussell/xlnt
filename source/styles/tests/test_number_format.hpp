#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include "pugixml.hpp"
#include <xlnt/xlnt.hpp>

class test_number_format : public CxxTest::TestSuite
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

    void test_simple_time()
    {
        auto time = xlnt::time(20, 15, 10);
        auto time_number = time.to_number();

        xlnt::number_format nf = xlnt::number_format::date_time2();
        auto formatted = nf.format(time_number, xlnt::calendar::windows_1900);

        TS_ASSERT_EQUALS(formatted, "8:15:10 PM");
    }

    void test_text_section_string()
    {
        xlnt::number_format nf;
        nf.set_format_string("General;General;General;[Green]\"a\"@\"b\"");

        auto formatted = nf.format("text");

        TS_ASSERT_EQUALS(formatted, "atextb");
    }

    void test_conditional_format()
    {
        xlnt::number_format nf;

        nf.set_format_string("[>5]\"first\"General;[>3]\"second\"General;\"third\"General");
        auto formatted = nf.format(6, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "first6");
        formatted = nf.format(4, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "second4");
        formatted = nf.format(5, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "second5");
        formatted = nf.format(3, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "third3");
        formatted = nf.format(2, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "third2");

        nf.set_format_string("[>=5]\"first\"General;[>=3]\"second\"General;\"third\"General");
        formatted = nf.format(5, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "first5");
        formatted = nf.format(6, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "first6");
        formatted = nf.format(4, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "second4");
        formatted = nf.format(3, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "second3");
        formatted = nf.format(2, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "third2");

        nf.set_format_string("[<1]\"first\"General;[<5]\"second\"General;\"third\"General");
        formatted = nf.format(0, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "first0");
        formatted = nf.format(1, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "second1");
        formatted = nf.format(5, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "third5");
        formatted = nf.format(6, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "third6");

        nf.set_format_string("[<=1]\"first\"General;[<=5]\"second\"General;\"third\"General");
        formatted = nf.format(-1000, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "first-1000");
        formatted = nf.format(0, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "first0");
        formatted = nf.format(1, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "first1");
        formatted = nf.format(4, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "second4");
        formatted = nf.format(5, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "second5");
        formatted = nf.format(6, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "third6");
        formatted = nf.format(1000, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "third1000");

        nf.set_format_string("[=1]\"first\"General;[=2]\"second\"General;\"third\"General");
        formatted = nf.format(1, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "first1");
        formatted = nf.format(2, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "second2");
        formatted = nf.format(3, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "third3");
        formatted = nf.format(0, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "third0");
    }
};
