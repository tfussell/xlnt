#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/writer/workbook_writer.hpp>

#include "helpers/path_helper.hpp"
#include "helpers/helper.hpp"

class test_theme : public CxxTest::TestSuite
{
public:
    void test_write_theme()
    {
        auto content = xlnt::write_theme();
        TS_ASSERT(Helper::EqualsFileContent(PathHelper::GetDataDirectory() + "/writer/expected/theme1.xml", content));
    }
};
