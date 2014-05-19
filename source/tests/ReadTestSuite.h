#pragma once

#include <fstream>
#include <iostream>
#include <cxxtest/TestSuite.h>

#include "../xlnt.h"
#include "PathHelper.h"

class ReadTestSuite : public CxxTest::TestSuite
{
public:
    ReadTestSuite()
    {
        xlnt::workbook wb_with_styles;
        auto path = PathHelper::GetDataDirectory() + "/genuine/empty-with-styles.xlsx";
        wb_with_styles.load(path);
        worksheet_with_styles = wb_with_styles.get_sheet_by_name("Sheet1");
        
        auto mac_wb_path = PathHelper::GetDataDirectory() + "/reader/date_1904.xlsx";
        mac_wb.load(mac_wb_path);
        mac_ws = mac_wb.get_sheet_by_name("Sheet1");
        
        auto win_wb_path = PathHelper::GetDataDirectory() + "/reader/date_1900.xlsx";
        win_wb.load(win_wb_path);
        win_ws = win_wb.get_sheet_by_name("Sheet1");
    }

    void test_read_standalone_worksheet()
    {
        auto path = PathHelper::GetDataDirectory() + "/reader/sheet2.xml";
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        {
            std::ifstream handle(path);
            ws = xlnt::reader::read_worksheet(handle, wb, "Sheet 2", {{1, "hello"}});
        }
        TS_ASSERT_DIFFERS(ws, nullptr);
        if(!(ws == nullptr))
        {
            TS_ASSERT_EQUALS(ws.cell("G5"), "hello");
            TS_ASSERT_EQUALS(ws.cell("D30"), 30);
            TS_ASSERT_EQUALS(ws.cell("K9"), 0.09);
        }
    }

    void test_read_standard_workbook()
    {
        auto path = PathHelper::GetDataDirectory() + "/genuine/empty.xlsx";
        xlnt::workbook wb;
        wb.load(path);
    }

    void test_read_standard_workbook_from_fileobj()
    {
        /*
        auto path = PathHelper::GetDataDirectory() + "/genuine/empty.xlsx";
        std::ifstream fo(path);
        xlnt::workbook wb;
        wb.load(fo);
         */
    }

    void test_read_worksheet()
    {
        auto path = PathHelper::GetDataDirectory() + "/genuine/empty.xlsx";
        xlnt::workbook wb;
        wb.load(path);
        auto sheet2 = wb.get_sheet_by_name("Sheet2 - Numbers");
        TS_ASSERT_DIFFERS(sheet2, nullptr);
        TS_ASSERT_EQUALS("This is cell G5", sheet2.cell("G5"));
        TS_ASSERT_EQUALS(18, sheet2.cell("D18"));
    }

    void test_read_nostring_workbook()
    {
        auto genuine_wb = PathHelper::GetDataDirectory() + "/genuine/empty-no-string.xlsx";
        xlnt::workbook wb;
        wb.load(genuine_wb);
    }

    void test_read_empty_file()
    {
        auto null_file = PathHelper::GetDataDirectory() + "/reader/null_file.xlsx";
        xlnt::workbook wb;
        TS_ASSERT_THROWS_ANYTHING(wb.load(null_file));
    }

        //@raises(InvalidFileException)
    void test_read_empty_archive()
    {
        auto null_file = PathHelper::GetDataDirectory() + "/reader/null_archive.xlsx";
        xlnt::workbook wb;
        TS_ASSERT_THROWS_ANYTHING(wb.load(null_file));
    }

    void test_read_dimension()
    {
        /*
        auto path = PathHelper::GetDataDirectory() + "/reader/sheet2.xml";
        std::ifstream handle(path);
        auto dimension = xlnt::reader::read_dimension(handle);
        TS_ASSERT_EQUALS({{"D", 1}, {"K", 30}}, dimension);
         */
    }

    void test_calculate_dimension_iter()
    {
        auto path = PathHelper::GetDataDirectory() + "/genuine/empty.xlsx";
        xlnt::workbook wb;
        wb.load(path);
        auto sheet2 = wb.get_sheet_by_name("Sheet2 - Numbers");
        auto dimensions = sheet2.calculate_dimension();
        TS_ASSERT_EQUALS("D1:K30", dimensions);
    }

    void test_get_highest_row_iter()
    {
        auto path = PathHelper::GetDataDirectory() + "/genuine/empty.xlsx";
        xlnt::workbook wb;
        wb.load(path);
        auto sheet2 = wb.get_sheet_by_name("Sheet2 - Numbers");
        auto max_row = sheet2.get_highest_row();
        TS_ASSERT_EQUALS(30, max_row);
    }

    void test_read_workbook_with_no_properties()
    {
        auto genuine_wb = PathHelper::GetDataDirectory() + "/genuine/empty_with_no_properties.xlsx";
        xlnt::workbook wb;
        wb.load(genuine_wb);
    }

    void test_read_general_style()
    {
        TS_ASSERT_EQUALS(worksheet_with_styles.cell("A1").get_style().get_number_format().get_format_code(), xlnt::number_format::format::general);
    }

    void test_read_date_style()
    {
        TS_ASSERT_EQUALS(worksheet_with_styles.cell("A2").get_style().get_number_format().get_format_code(), xlnt::number_format::format::date_xlsx14);
    }

    void test_read_number_style()
    {
        TS_ASSERT_EQUALS(worksheet_with_styles.cell("A3").get_style().get_number_format().get_format_code(), xlnt::number_format::format::number00);
    }

    void test_read_time_style()
    {
        TS_ASSERT_EQUALS(worksheet_with_styles.cell("A4").get_style().get_number_format().get_format_code(), xlnt::number_format::format::date_time3);
    }

    void test_read_percentage_style()
    {
        TS_ASSERT_EQUALS(worksheet_with_styles.cell("A5").get_style().get_number_format().get_format_code(), xlnt::number_format::format::percentage00)
    }

    void test_read_win_base_date()
    {
        //TS_ASSERT_EQUALS(win_wb.get_properties().get_excel_base_date(), CALENDAR_WINDOWS_1900);
    }

    void test_read_mac_base_date()
    {
        //TS_ASSERT_EQUALS(mac_wb.get_properties().get_excel_base_date(), CALENDAR_MAC_1904);
    }

    void test_read_date_style_mac()
    {
        TS_ASSERT_EQUALS(mac_ws.cell("A1").get_style().get_number_format().get_format_code(), xlnt::number_format::format::date_xlsx14);
    }

    void test_read_date_style_win()
    {
        TS_ASSERT_EQUALS(win_ws.cell("A1").get_style().get_number_format().get_format_code(), xlnt::number_format::format::date_xlsx14);
    }

    void test_read_date_value()
    {
        /*
        auto datetuple = (2011, 10, 31);
        auto dt = datetime(datetuple[0], datetuple[1], datetuple[2]);
        TS_ASSERT_EQUALS(mac_ws.cell("A1"), dt);
        TS_ASSERT_EQUALS(win_ws.cell("A1"), dt);
        TS_ASSERT_EQUALS(mac_ws.cell("A1"), win_ws.cell("A1"));
         */
    }
    
private:
    xlnt::worksheet worksheet_with_styles;
    xlnt::workbook mac_wb;
    xlnt::worksheet mac_ws;
    xlnt::workbook win_wb;
    xlnt::worksheet win_ws;
};
