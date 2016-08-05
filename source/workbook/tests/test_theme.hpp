#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <helpers/path_helper.hpp>
#include <helpers/xml_helper.hpp>
#include <xlnt/workbook/workbook.hpp>

class test_theme : public CxxTest::TestSuite
{
public:
    void test_write_theme()
    {
        xlnt::workbook wb;
		wb.set_theme(xlnt::theme());

        TS_ASSERT(xml_helper::file_matches_workbook_part(
			path_helper::get_data_directory("writer/expected/theme1.xml"),
			wb, xlnt::constants::part_theme()));
    }
};
