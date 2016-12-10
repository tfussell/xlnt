#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <helpers/xml_helper.hpp>

class test_tests : public CxxTest::TestSuite
{
public:
    void test_compare()
    {
        TS_ASSERT(!xml_helper::compare_xml_exact("<a/>", "<b/>", true));
        TS_ASSERT(!xml_helper::compare_xml_exact("<a/>", "<a b=\"4\"/>", true));
        TS_ASSERT(!xml_helper::compare_xml_exact("<a b=\"3\"/>", "<a/>", true));
        TS_ASSERT(!xml_helper::compare_xml_exact("<a c=\"4\"/>", "<a b=\"4\"/>", true));
        TS_ASSERT(!xml_helper::compare_xml_exact("<a b=\"3\"/>", "<a b=\"4\"/>", true));
        TS_ASSERT(!xml_helper::compare_xml_exact("<a>text</a>", "<a>txet</a>", true));
        TS_ASSERT(!xml_helper::compare_xml_exact("<a>text</a>", "<a><b>txet</b></a>", true));
        TS_ASSERT(xml_helper::compare_xml_exact("<a/>", "<a> </a>", true));
        TS_ASSERT(xml_helper::compare_xml_exact("<a b=\"3\"/>", "<a b=\"3\"></a>", true));
        TS_ASSERT(xml_helper::compare_xml_exact("<a>text</a>", "<a>text</a>", true));
    }
};
