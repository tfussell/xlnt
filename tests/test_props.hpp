#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/xlnt.hpp>
#include "helpers/path_helper.hpp"
#include "helpers/helper.hpp"

class test_props : public CxxTest::TestSuite
{
public:
    void test_read_properties_core()
    {
        xlnt::zip_file archive(PathHelper::GetDataDirectory() + "/genuine/empty.xlsx", xlnt::file_mode::open);
        auto content = archive.get_file_contents("docProps/core.xml");
        auto prop = xlnt::reader::read_properties_core(content);
        TS_ASSERT_EQUALS(prop.creator, "*.*");
        TS_ASSERT_EQUALS(prop.last_modified_by, "Charlie Clark");
        TS_ASSERT_EQUALS(prop.created, xlnt::datetime(2010, 4, 9, 20, 43, 12));
        TS_ASSERT_EQUALS(prop.modified, xlnt::datetime(2014, 1, 2, 14, 53, 6));
    }

    void test_read_sheets_titles()
    {
        xlnt::zip_file archive(PathHelper::GetDataDirectory() + "/genuine/empty.xlsx", xlnt::file_mode::open);
        auto content = archive.get_file_contents("docProps/core.xml");
        std::vector<std::string> expected_titles = {"Sheet1 - Text", "Sheet2 - Numbers", "Sheet3 - Formulas", "Sheet4 - Dates"};
        int i = 0;
        for(auto sheet : xlnt::reader::read_sheets(archive))
        {
            TS_ASSERT_EQUALS(sheet.second, expected_titles[i++]);
        }
    }

    void test_read_properties_core_libre()
    {
        xlnt::zip_file archive(PathHelper::GetDataDirectory() + "/genuine/empty_libre.xlsx", xlnt::file_mode::open);
        auto content = archive.get_file_contents("docProps/core.xml");
        auto prop = xlnt::reader::read_properties_core(content);
        TS_ASSERT_EQUALS(prop.excel_base_date, xlnt::calendar::windows_1900);
    }

    void test_read_sheets_titles_libre()
    {
        xlnt::zip_file archive(PathHelper::GetDataDirectory() + "/genuine/empty_libre.xlsx", xlnt::file_mode::open);
        auto content = archive.get_file_contents("docProps/core.xml");
        std::vector<std::string> expected_titles = {"Sheet1 - Text", "Sheet2 - Numbers", "Sheet3 - Formulas", "Sheet4 - Dates"};
        int i = 0;
        for(auto sheet : xlnt::reader::read_sheets(archive))
        {
            TS_ASSERT_EQUALS(sheet.second, expected_titles[i++]);
        }
    }

    void test_write_properties_core()
    {
        xlnt::document_properties prop;
        prop.creator = "TEST_USER";
        prop.last_modified_by = "SOMEBODY";
        prop.created = xlnt::datetime(2010, 4, 1, 20, 30, 00);
        prop.modified = xlnt::datetime(2010, 4, 5, 14, 5, 30);
        auto content = xlnt::writer::write_properties_core(prop);
        TS_ASSERT(Helper::EqualsFileContent(PathHelper::GetDataDirectory() + "/writer/expected/core.xml", content));
    }

    void test_write_properties_app()
    {
        xlnt::workbook wb;
        wb.create_sheet();
        wb.create_sheet();
        auto content = xlnt::writer::write_properties_app(wb);
        TS_ASSERT(Helper::EqualsFileContent(PathHelper::GetDataDirectory() + "/writer/expected/app.xml", content));
    }
};
