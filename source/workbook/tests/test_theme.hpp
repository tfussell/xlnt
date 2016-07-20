#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <detail/theme_serializer.hpp>
#include <helpers/path_helper.hpp>
#include <helpers/xml_helper.hpp>
#include <xlnt/workbook/workbook.hpp>

class test_theme : public CxxTest::TestSuite
{
public:
    void test_write_theme()
    {
        xlnt::workbook wb;
        xlnt::theme_serializer serializer;
        pugi::xml_document xml;
        serializer.write_theme(wb.get_loaded_theme(), xml);
        TS_ASSERT(xml_helper::compare_xml(path_helper::get_data_directory() + "/writer/expected/theme1.xml", xml));
    }
};
