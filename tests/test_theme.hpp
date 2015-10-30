#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/s11n/theme_serializer.hpp>
#include <xlnt/workbook/workbook.hpp>

#include "helpers/path_helper.hpp"
#include "helpers/helper.hpp"

class test_theme : public CxxTest::TestSuite
{
public:
    void test_write_theme()
    {
        xlnt::theme_serializer serializer;
        xlnt::workbook wb;
        auto content = serializer.write_theme(wb.get_loaded_theme());
        TS_ASSERT(Helper::compare_xml(PathHelper::GetDataDirectory() + "/writer/expected/theme1.xml", content));
    }
};
