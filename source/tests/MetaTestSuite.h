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
        wb = Workbook();
        wb.create_sheet();
        wb.create_sheet();
        content = write_content_types(wb);
        reference_file = os.path.join(DATADIR, "writer", "expected",
            "[Content_Types].xml");
        assert_equals_file_content(reference_file, content);
    }

    void test_write_root_rels()
    {
        wb = Workbook();
        content = write_root_rels(wb);
        reference_file = os.path.join(DATADIR, "writer", "expected", ".rels");
        assert_equals_file_content(reference_file, content);
    }
};
