#pragma once

#include <cxxtest/TestSuite.h>

#include "../xlnt.h"
#include "pugixml.hpp"
#include "PathHelper.h"

class ZipFileTestSuite : public CxxTest::TestSuite
{
public:
    ZipFileTestSuite()
    {
        existing_zip = PathHelper::GetDataDirectory() + "/reader/a.zip";
        existing_xlsx = PathHelper::GetDataDirectory() + "/reader/date_1900.xlsx";
        new_xlsx = PathHelper::GetDataDirectory() + "/writer/new.xlsx";
    }

    void test_existing_package()
    {
        xlnt::zip_file package(existing_xlsx, xlnt::file_mode::open, xlnt::file_access::read);
    }

    void test_new_package()
    {
        xlnt::zip_file package(new_xlsx, xlnt::file_mode::create, xlnt::file_access::read_write);
    }

    void test_read_text()
    {
        xlnt::zip_file package(existing_xlsx, xlnt::file_mode::open, xlnt::file_access::read);
        auto contents = package.get_file_contents("xl/_rels/workbook.xml.rels");
        auto expected = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet3.xml\"/>"
            "<Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>"
            "<Relationship Id=\"rId5\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>"
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
            "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet2.xml\"/>"
            "</Relationships>";
        TS_ASSERT_EQUALS(contents, expected);
    }

    void test_write_text()
    {
        {
            xlnt::zip_file package(new_xlsx, xlnt::file_mode::create, xlnt::file_access::write);
            package.set_file_contents("a.txt", "something else");
        }

        {
            xlnt::zip_file package(new_xlsx, xlnt::file_mode::open, xlnt::file_access::read);
            auto contents = package.get_file_contents("a.txt");
            TS_ASSERT_EQUALS(contents, "something else");
        }
    }

    void test_read_xml()
    {
        xlnt::zip_file package(existing_zip, xlnt::file_mode::open, xlnt::file_access::read);
        auto contents = package.get_file_contents("a.xml");

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
    std::string existing_zip;
    std::string existing_xlsx;
    std::string new_xlsx;
};
