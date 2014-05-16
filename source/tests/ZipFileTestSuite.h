#pragma once

#include <cxxtest/TestSuite.h>

#include "../xlnt.h"
#include "pugixml.hpp"

class ZipFileTestSuite : public CxxTest::TestSuite
{
public:
    ZipFileTestSuite()
    {
        existing_xlsx = "/Users/thomas/Development/xlnt/source/tests/test_data/reader/date_1990.xlsx";
        new_xlsx = "/Users/thomas/Development/xlnt/source/tests/test_data/writer/new.xlsx";
    }

    void test_existing_package()
    {
        xlnt::zip_file package(existing_xlsx, xlnt::file_mode::open, xlnt::file_access::read);
    }

    void test_new_package()
    {
        {
            xlnt::zip_file package(new_xlsx, xlnt::file_mode::create, xlnt::file_access::read_write);
            package.flush();
        }
    }

    void test_read_text()
    {
        xlnt::zip_file package(existing_xlsx, xlnt::file_mode::open, xlnt::file_access::read);
        auto contents = package.get_file_contents("a.txt");
        TS_ASSERT_EQUALS(contents, "a.txt");
    }

    void test_write_text()
    {
        {
            xlnt::zip_file package(new_xlsx, xlnt::file_mode::open, xlnt::file_access::read);
            package.set_file_contents("a.txt", "something else");
        }

        {
            xlnt::zip_file package(new_xlsx, xlnt::file_mode::open, xlnt::file_access::write);
            auto contents = package.get_file_contents("a.txt");
            TS_ASSERT_EQUALS(contents, "something else");
        }
    }

    void test_read_xml()
    {
        xlnt::zip_file package(existing_xlsx, xlnt::file_mode::open, xlnt::file_access::read);
        auto contents = package.get_file_contents("a.txt");

        pugi::xml_document doc;
        doc.load(contents.c_str());

        auto root_element = doc.child("root");
        TS_ASSERT_DIFFERS(root_element, nullptr)

        auto child_element = root_element.child("child");
        TS_ASSERT_DIFFERS(child_element, nullptr)

        TS_ASSERT_EQUALS(std::string(child_element.attribute("attribute").as_string()), "attribute")

        auto element_element = root_element.child("element");
        TS_ASSERT_DIFFERS(element_element, nullptr)

        TS_ASSERT_EQUALS(std::string(element_element.text().as_string()), "Text")
    }

private:
    std::string existing_xlsx;
    std::string new_xlsx;
};
