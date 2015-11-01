#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/serialization/theme_serializer.hpp>
#include <xlnt/workbook/workbook.hpp>

#include "helpers/path_helper.hpp"
#include "helpers/helper.hpp"

class test_theme : public CxxTest::TestSuite
{
public:
    void test_write_theme()
    {
        xlnt::workbook wb;
        xlnt::theme_serializer serializer;
        auto xml = serializer.write_theme(wb.get_loaded_theme());
        TS_ASSERT(Helper::compare_xml(PathHelper::GetDataDirectory() + "/writer/expected/theme1.xml", xml));
    }
};
