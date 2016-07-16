#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/packaging/app_properties.hpp>
#include <detail/workbook_serializer.hpp>

#include <helpers/path_helper.hpp>
#include <helpers/helper.hpp>

class test_core : public CxxTest::TestSuite
{
public:
    void test_read_properties_core()
    {
        auto path = PathHelper::GetDataDirectory() + "/genuine/empty.xlsx";
        xlnt::workbook wb;
        wb.load(path);
        auto &prop = wb.get_properties();
        TS_ASSERT_EQUALS(prop.creator, "*.*");
        TS_ASSERT_EQUALS(prop.last_modified_by, "Charlie Clark");
        TS_ASSERT_EQUALS(prop.created, xlnt::datetime(2010, 4, 9, 20, 43, 12));
        TS_ASSERT_EQUALS(prop.modified, xlnt::datetime(2014, 1, 2, 14, 53, 6));
    }

    void test_read_sheets_titles()
    {
        auto path = PathHelper::GetDataDirectory() + "/genuine/empty.xlsx";
        
        const std::vector<std::string> expected_titles = {"Sheet1 - Text", "Sheet2 - Numbers", "Sheet3 - Formulas", "Sheet4 - Dates"};
        
        std::size_t i = 0;
        
        xlnt::workbook wb;
        wb.load(path);
        
        for(auto sheet : wb)
        {
            TS_ASSERT_EQUALS(sheet.get_title(), expected_titles.at(i++));
        }
    }

    void test_read_properties_core_libre()
    {
        xlnt::zip_file archive(PathHelper::GetDataDirectory() + "/genuine/empty_libre.xlsx");
        auto content = archive.read("docProps/core.xml");
        xlnt::workbook wb;
        xlnt::workbook_serializer serializer(wb);
        pugi::xml_document xml;
        serializer.read_properties_core(xml);
        auto &prop = wb.get_properties();
        TS_ASSERT_EQUALS(prop.excel_base_date, xlnt::calendar::windows_1900);
    }

    void test_read_sheets_titles_libre()
    {
        auto path = PathHelper::GetDataDirectory() + "/genuine/empty_libre.xlsx";
        
        const std::vector<std::string> expected_titles = {"Sheet1 - Text", "Sheet2 - Numbers", "Sheet3 - Formulas", "Sheet4 - Dates"};
        
        std::size_t i = 0;
        
        xlnt::workbook wb;
        wb.load(path);
        
        for(auto sheet : wb)
        {
            TS_ASSERT_EQUALS(sheet.get_title(), expected_titles.at(i++));
        }
    }

    void test_write_properties_core()
    {
        xlnt::workbook wb;
        xlnt::document_properties &prop = wb.get_properties();
        prop.creator = "TEST_USER";
        prop.last_modified_by = "SOMEBODY";
        prop.created = xlnt::datetime(2010, 4, 1, 20, 30, 00);
        prop.modified = xlnt::datetime(2010, 4, 5, 14, 5, 30);
        xlnt::workbook_serializer serializer(wb);
        pugi::xml_document xml;
        serializer.write_properties_core(xml);
        TS_ASSERT(Helper::compare_xml(PathHelper::GetDataDirectory() + "/writer/expected/core.xml", xml));
    }

    void test_write_properties_app()
    {
        xlnt::workbook wb;
        wb.get_app_properties().application = "Microsoft Excel";
        wb.get_app_properties().app_version = "12.0000";
        wb.get_app_properties().company = "Company";
        wb.create_sheet();
        wb.create_sheet();
        xlnt::workbook_serializer serializer(wb);
        pugi::xml_document xml;
        serializer.write_properties_app(xml);
        TS_ASSERT(Helper::compare_xml(PathHelper::GetDataDirectory() + "/writer/expected/app.xml", xml));
    }
};
