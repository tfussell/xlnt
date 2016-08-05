#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <detail/constants.hpp>
#include <helpers/path_helper.hpp>
#include <helpers/xml_helper.hpp>
#include <xlnt/workbook/workbook.hpp>

class test_core : public CxxTest::TestSuite
{
public:
    void test_read_properties_core()
    {
        xlnt::workbook wb;
        wb.load(path_helper::get_data_directory("genuine/empty.xlsx"));

        TS_ASSERT_EQUALS(wb.get_creator(), "*.*");
        TS_ASSERT_EQUALS(wb.get_last_modified_by(), "Charlie Clark");
        TS_ASSERT_EQUALS(wb.get_created(), xlnt::datetime(2010, 4, 9, 20, 43, 12));
        TS_ASSERT_EQUALS(wb.get_modified(), xlnt::datetime(2014, 1, 2, 14, 53, 6));
    }

    void test_read_sheets_titles()
    {
        const std::vector<std::string> expected_titles = {"Sheet1 - Text", "Sheet2 - Numbers", "Sheet3 - Formulas", "Sheet4 - Dates"};
        
        xlnt::workbook wb;
        wb.load(path_helper::get_data_directory("genuine/empty.xlsx"));
        
		TS_ASSERT_EQUALS(wb.get_sheet_titles(), expected_titles);
    }

    void test_read_properties_core_libre()
    {
		xlnt::workbook wb;
		wb.load(path_helper::get_data_directory("genuine/empty_libre.xlsx"));
        TS_ASSERT_EQUALS(wb.get_base_date(), xlnt::calendar::windows_1900);
    }

    void test_read_sheets_titles_libre()
    {
        const std::vector<std::string> expected_titles = {"Sheet1 - Text", "Sheet2 - Numbers", "Sheet3 - Formulas", "Sheet4 - Dates"};
        
        xlnt::workbook wb;
        wb.load(path_helper::get_data_directory("genuine/empty_libre.xlsx"));
		auto title_iter = expected_titles.begin();
        
        for(auto ws : wb)
        {
            TS_ASSERT_EQUALS(ws.get_title(), *(title_iter++));
        }
    }

    void test_write_properties_core()
    {
        xlnt::workbook wb;
        wb.set_creator("TEST_USER");
        wb.set_last_modified_by("SOMEBODY");
        wb.set_created(xlnt::datetime(2010, 4, 1, 20, 30, 00));
        wb.set_modified(xlnt::datetime(2010, 4, 5, 14, 5, 30));

        TS_ASSERT(xml_helper::file_matches_workbook_part(
			path_helper::get_data_directory("writer/expected/core.xml"), 
			wb, xlnt::constants::part_core()));
    }

    void test_write_properties_app()
    {
        xlnt::workbook wb;
        wb.set_application("Microsoft Excel");
        wb.set_app_version("12.0000");
        wb.set_company("Company");
        wb.create_sheet();
        wb.create_sheet();

		TS_ASSERT(xml_helper::file_matches_workbook_part(
			path_helper::get_data_directory("writer/expected/core.xml"),
			wb, xlnt::constants::part_core()));
    }
};
