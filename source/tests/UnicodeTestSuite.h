#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include "../xlnt.h"

class UnicodeTestSuite : public CxxTest::TestSuite
{
public:
    UnicodeTestSuite()
    {

    }

    void test_read_workbook_with_unicode_character()
    {
        //unicode_wb = os.path.join(DATADIR, "genuine", "unicode.xlsx");
        //wb = load_workbook(filename = unicode_wb);
    }
};
