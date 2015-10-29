#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/s11n/workbook_serializer.hpp>

#include "helpers/path_helper.hpp"
#include "helpers/helper.hpp"

class test_props : public CxxTest::TestSuite
{
public:
    void test_read_properties_core()
    {
        xlnt::zip_file archive(PathHelper::GetDataDirectory() + "/genuine/empty.xlsx");
        auto content = archive.read("docProps/core.xml");
        xlnt::workbook wb;
        xlnt::workbook_serializer serializer(wb);
        auto prop = serializer.read_properties_core(content);
        TS_ASSERT_EQUALS(prop.creator, "*.*");
        TS_ASSERT_EQUALS(prop.last_modified_by, "Charlie Clark");
        TS_ASSERT_EQUALS(prop.created, xlnt::datetime(2010, 4, 9, 20, 43, 12));
        TS_ASSERT_EQUALS(prop.modified, xlnt::datetime(2014, 1, 2, 14, 53, 6));
    }

    void test_read_sheets_titles()
    {
        xlnt::zip_file archive(PathHelper::GetDataDirectory() + "/genuine/empty.xlsx");
        auto content = archive.read("docProps/core.xml");
        
        const std::vector<std::string> expected_titles = {"Sheet1 - Text", "Sheet2 - Numbers", "Sheet3 - Formulas", "Sheet4 - Dates"};
        
        std::size_t i = 0;
        for(auto sheet : xlnt::read_sheets(archive))
        {
            TS_ASSERT_EQUALS(sheet.second, expected_titles.at(i++));
        }
    }

    void test_read_properties_core_libre()
    {
        xlnt::zip_file archive(PathHelper::GetDataDirectory() + "/genuine/empty_libre.xlsx");
        auto content = archive.read("docProps/core.xml");
        auto prop = xlnt::read_properties_core(content);
        TS_ASSERT_EQUALS(prop.excel_base_date, xlnt::calendar::windows_1900);
    }

    void test_read_sheets_titles_libre()
    {
        xlnt::zip_file archive(PathHelper::GetDataDirectory() + "/genuine/empty_libre.xlsx");
        auto content = archive.read("docProps/core.xml");
        
        const std::vector<std::string> expected_titles = {"Sheet1 - Text", "Sheet2 - Numbers", "Sheet3 - Formulas", "Sheet4 - Dates"};
        
        std::size_t i = 0;
        for(auto sheet : xlnt::read_sheets(archive))
        {
            TS_ASSERT_EQUALS(sheet.second, expected_titles.at(i++));
        }
    }

    void test_write_properties_core()
    {
        xlnt::document_properties prop;
        prop.creator = "TEST_USER";
        prop.last_modified_by = "SOMEBODY";
        prop.created = xlnt::datetime(2010, 4, 1, 20, 30, 00);
        prop.modified = xlnt::datetime(2010, 4, 5, 14, 5, 30);
        auto content = xlnt::write_properties_core(prop);
        TS_ASSERT(Helper::EqualsFileContent(PathHelper::GetDataDirectory() + "/writer/expected/core.xml", content));
    }

    void test_write_properties_app()
    {
        xlnt::workbook wb;
        wb.create_sheet();
        wb.create_sheet();
        auto content = xlnt::write_properties_app(wb);
        TS_ASSERT(Helper::EqualsFileContent(PathHelper::GetDataDirectory() + "/writer/expected/app.xml", content));
    }
};
