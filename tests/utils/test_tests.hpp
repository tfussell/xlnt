#pragma once

#include <iostream>
#include <helpers/test_suite.hpp>

#include <helpers/xml_helper.hpp>

class test_tests : public test_suite
{
public:
    void test_compare()
    {
        assert(!xml_helper::compare_xml_exact("<a/>", "<b/>", true));
        assert(!xml_helper::compare_xml_exact("<a/>", "<a b=\"4\"/>", true));
        assert(!xml_helper::compare_xml_exact("<a b=\"3\"/>", "<a/>", true));
        assert(!xml_helper::compare_xml_exact("<a c=\"4\"/>", "<a b=\"4\"/>", true));
        assert(!xml_helper::compare_xml_exact("<a b=\"3\"/>", "<a b=\"4\"/>", true));
        assert(!xml_helper::compare_xml_exact("<a>text</a>", "<a>txet</a>", true));
        assert(!xml_helper::compare_xml_exact("<a>text</a>", "<a><b>txet</b></a>", true));
        assert(xml_helper::compare_xml_exact("<a/>", "<a> </a>", true));
        assert(xml_helper::compare_xml_exact("<a b=\"3\"/>", "<a b=\"3\"></a>", true));
        assert(xml_helper::compare_xml_exact("<a>text</a>", "<a>text</a>", true));
    }
};
