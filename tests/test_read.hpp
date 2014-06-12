#pragma once

#include <fstream>
#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/xlnt.hpp>
#include "helpers/path_helper.hpp"

class test_read : public CxxTest::TestSuite
{
public:
    void test_read_standalone_worksheet()
    {
        auto path = PathHelper::GetDataDirectory() + "/reader/sheet2.xml";
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        {
            std::ifstream handle(path);
            ws = xlnt::reader::read_worksheet(handle, wb, "Sheet 2", {"hello"});
        }
        TS_ASSERT_DIFFERS(ws, nullptr);
        if(!(ws == nullptr))
        {
            TS_ASSERT_EQUALS(ws.get_cell("G5"), "hello");
            TS_ASSERT_EQUALS(ws.get_cell("D30"), 30);
            TS_ASSERT_EQUALS(ws.get_cell("K9"), 0.09);
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
        auto path = PathHelper::GetDataDirectory() + "/genuine/empty.xlsx";
        std::ifstream fo(path);
        xlnt::workbook wb;
        wb.load(fo);
    }

    void test_read_worksheet()
    {
        auto path = PathHelper::GetDataDirectory() + "/genuine/empty.xlsx";
        xlnt::workbook wb;
        wb.load(path);
        auto sheet2 = wb.get_sheet_by_name("Sheet2 - Numbers");
        TS_ASSERT_DIFFERS(sheet2, nullptr);
        TS_ASSERT_EQUALS("This is cell G5", sheet2.get_cell("G5"));
        TS_ASSERT_EQUALS(18, sheet2.get_cell("D18"));
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
        auto path = PathHelper::GetDataDirectory() + "/reader/sheet2.xml";
        std::ifstream file(path);
        std::stringstream ss;
        ss << file.rdbuf();
        auto dimension = xlnt::reader::read_dimension(ss.str());
        TS_ASSERT_EQUALS("D1:AA30", dimension);
    }

    void test_calculate_dimension_iter()
    {
        auto path = PathHelper::GetDataDirectory() + "/genuine/empty.xlsx";
        xlnt::workbook wb;
        wb.load(path);
        auto sheet2 = wb.get_sheet_by_name("Sheet2 - Numbers");
        auto dimensions = sheet2.calculate_dimension();
        TS_ASSERT_EQUALS("D1:AA30", dimensions.to_string());
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
        //TS_ASSERT_EQUALS(worksheet_with_styles.get_cell("A1").get_style().get_number_format().get_format_code(), xlnt::number_format::format::general);
    }

    void test_read_date_style()
    {
        //TS_ASSERT_EQUALS(worksheet_with_styles.get_cell("A2").get_style().get_number_format().get_format_code(), xlnt::number_format::format::date_xlsx14);
    }

    void test_read_number_style()
    {
        //TS_ASSERT_EQUALS(worksheet_with_styles.get_cell("A3").get_style().get_number_format().get_format_code(), xlnt::number_format::format::number00);
    }

    void test_read_time_style()
    {
        //TS_ASSERT_EQUALS(worksheet_with_styles.get_cell("A4").get_style().get_number_format().get_format_code(), xlnt::number_format::format::date_time3);
    }

    void test_read_percentage_style()
    {
        //TS_ASSERT_EQUALS(worksheet_with_styles.get_cell("A5").get_style().get_number_format().get_format_code(), xlnt::number_format::format::percentage00)
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
        //TS_ASSERT_EQUALS(mac_ws.get_cell("A1").get_style().get_number_format().get_format_code(), xlnt::number_format::format::date_xlsx14);
    }

    void test_read_date_style_win()
    {
        //TS_ASSERT_EQUALS(win_ws.get_cell("A1").get_style().get_number_format().get_format_code(), xlnt::number_format::format::date_xlsx14);
    }

    void test_read_date_value()
    {
        //auto path = PathHelper::GetDataDirectory() + "/genuine/empty-with-styles.xlsx";
        //wb_with_styles.load(path);
        //worksheet_with_styles = wb_with_styles.get_sheet_by_name("Sheet1");
        
        auto mac_wb_path = PathHelper::GetDataDirectory() + "/reader/date_1904.xlsx";
        xlnt::workbook mac_wb;
        mac_wb.load(mac_wb_path);
        auto mac_ws = mac_wb.get_sheet_by_name("Sheet1");
        
        auto win_wb_path = PathHelper::GetDataDirectory() + "/reader/date_1900.xlsx";
        xlnt::workbook win_wb;
        win_wb.load(win_wb_path);
        auto win_ws = win_wb.get_sheet_by_name("Sheet1");
        
        xlnt::datetime dt(2011, 10, 31);
        TS_ASSERT_EQUALS(mac_ws.get_cell("A1"), dt);
        TS_ASSERT_EQUALS(win_ws.get_cell("A1"), dt);
        TS_ASSERT_EQUALS(mac_ws.get_cell("A1"), win_ws.get_cell("A1"));
    }
};
