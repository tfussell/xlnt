#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include "../xlnt.h"

class MetaTestSuite : public CxxTest::TestSuite
{
public:
    MetaTestSuite()
    {

    }

    void test_write_content_types()
    {
        xlnt::workbook wb;
        wb.create_sheet();
        wb.create_sheet();
        auto content = xlnt::workbook::write_content_types(wb);
        std::string reference_file = DATADIR + "/writer/expected/[Content_Types].xml";
        assert_equals_file_content(reference_file, content);
    }

    void test_write_root_rels()
    {
        xlnt::workbook wb;
        wb.create_sheet();
        auto content = xlnt::workbook::write_root_rels(wb);
        std::string reference_file = DATADIR + "/writer/expected/.rels";
        assert_equals_file_content(reference_file, content);
    }
};
