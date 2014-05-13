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

    bool equals_file_content(const std::string &file1, const std::string &file2)
    {
        return false;
    }

    void test_write_content_types()
    {
        xlnt::workbook wb;
        wb.create_sheet();
        wb.create_sheet();
        auto content = xlnt::writer::write_content_types(wb);
        std::string reference_file = DATADIR + "/writer/expected/[Content_Types].xml";
        TS_ASSERT(equals_file_content(reference_file, content));
    }

    void test_write_root_rels()
    {
        xlnt::workbook wb;
        wb.create_sheet();
        auto content = xlnt::writer::write_root_rels(wb);
        std::string reference_file = DATADIR + "/writer/expected/.rels";
        TS_ASSERT(equals_file_content(reference_file, content));
    }

private:
    const std::string DATADIR = "../../source/tests";
};
