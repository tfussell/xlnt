#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/xlnt.h>

class PropsTestSuite : public CxxTest::TestSuite
{
public:
    PropsTestSuite()
    {

    }

    class TestReaderProps
    {
        void setup_class(int cls)
        {
            //cls.genuine_filename = os.path.join(DATADIR, "genuine", "empty.xlsx");
            //cls.archive = ZipFile(cls.genuine_filename, "r", ZIP_DEFLATED);
        }

        void teardown_class(int cls)
        {
            //cls.archive.close();
        }
    };

    void test_read_properties_core()
    {
        //content = archive.read(ARC_CORE)
        //    prop = read_properties_core(content)
        //    TS_ASSERT_EQUALS(prop.creator, "*.*")
        //    eacute = chr(233)
        //    TS_ASSERT_EQUALS(prop.last_modified_by, "Aur" + eacute + "lien Camp" + eacute + "as")
        //    TS_ASSERT_EQUALS(prop.created, datetime(2010, 4, 9, 20, 43, 12))
        //    TS_ASSERT_EQUALS(prop.modified, datetime(2011, 2, 9, 13, 49, 32))
    }

    void test_read_sheets_titles()
    {
        //content = archive.read(ARC_WORKBOOK);
        //sheet_titles = read_sheets_titles(content);
        //TS_ASSERT_EQUALS(sheet_titles, \
        //    ["Sheet1 - Text", "Sheet2 - Numbers", "Sheet3 - Formulas", "Sheet4 - Dates"]);
    }

         // Just tests that the correct date / time format is returned from LibreOffice saved version

        void setup_class(int cls)
        {
            //cls.genuine_filename = os.path.join(DATADIR, "genuine", "empty_libre.xlsx")
            //    cls.archive = ZipFile(cls.genuine_filename, "r", ZIP_DEFLATED)
        }

        void teardown_class(int cls)
        {
            //cls.archive.close()
        }

        void test_read_properties_core2()
        {
        //    content = archive.read(ARC_CORE)
        //        prop = read_properties_core(content)
        //        TS_ASSERT_EQUALS(prop.excel_base_date, CALENDAR_WINDOWS_1900)
        }

        void test_read_sheets_titles2()
        {
            //content = archive.read(ARC_WORKBOOK)
            //    sheet_titles = read_sheets_titles(content)
            //    TS_ASSERT_EQUALS(sheet_titles, \
            //    ["Sheet1 - Text", "Sheet2 - Numbers", "Sheet3 - Formulas", "Sheet4 - Dates"])
        }

    void test_write_properties_core()
    {
        //prop.creator = "TEST_USER"
        //    prop.last_modified_by = "SOMEBODY"
        //    prop.created = datetime(2010, 4, 1, 20, 30, 00)
        //    prop.modified = datetime(2010, 4, 5, 14, 5, 30)
        //    content = write_properties_core(prop)
        //    assert_equals_file_content(
        //    os.path.join(DATADIR, "writer", "expected", "core.xml"),
        //    content)
    }

    void test_write_properties_app()
    {
        //wb = Workbook()
        //    wb.create_sheet()
        //    wb.create_sheet()
        //    content = write_properties_app(wb)
        //    assert_equals_file_content(
        //    os.path.join(DATADIR, "writer", "expected", "app.xml"),
        //    content)
    }
};
