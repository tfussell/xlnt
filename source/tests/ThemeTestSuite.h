#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include "../xlnt.h"

class ThemeTestSuite : public CxxTest::TestSuite
{
public:
    ThemeTestSuite()
    {

    }

    void test_write_theme()
    {
        //content = write_theme();
        //assert_equals_file_content(
        //    os.path.join(DATADIR, "writer", "expected", "theme1.xml"), content);
    }
};
