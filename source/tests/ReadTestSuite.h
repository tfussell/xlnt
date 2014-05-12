#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include "../xlnt.h"

class ReadTestSuite : public CxxTest::TestSuite
{
public:
    ReadTestSuite()
    {

    }

    void test_read_standalone_worksheet()
    {
        path = os.path.join(DATADIR, "reader", "sheet2.xml")
            ws = None
            handle = open(path)
            try :
            ws = read_worksheet(handle.read(), DummyWb(),
            "Sheet 2", {1: "hello"}, {1: Style()})
            finally :
            handle.close()
            assert isinstance(ws, Worksheet)
            TS_ASSERT_EQUALS(ws.cell("G5").value, "hello")
            TS_ASSERT_EQUALS(ws.cell("D30").value, 30)
            TS_ASSERT_EQUALS(ws.cell("K9").value, 0.09)
    }

    void test_read_standard_workbook()
    {
        path = os.path.join(DATADIR, "genuine", "empty.xlsx")
            wb = load_workbook(path)
            assert isinstance(wb, Workbook)
    }

    void test_read_standard_workbook_from_fileobj()
    {
        path = os.path.join(DATADIR, "genuine", "empty.xlsx")
            fo = open(path, mode = "rb")
            wb = load_workbook(fo)
            assert isinstance(wb, Workbook)
    }

    void test_read_worksheet()
    {
        path = os.path.join(DATADIR, "genuine", "empty.xlsx")
            wb = load_workbook(path)
            sheet2 = wb.get_sheet_by_name("Sheet2 - Numbers")
            assert isinstance(sheet2, Worksheet)
            TS_ASSERT_EQUALS("This is cell G5", sheet2.cell("G5").value)
            TS_ASSERT_EQUALS(18, sheet2.cell("D18").value)
    }

    void test_read_nostring_workbook()
    {
        genuine_wb = os.path.join(DATADIR, "genuine", "empty-no-string.xlsx")
            wb = load_workbook(genuine_wb)
            assert isinstance(wb, Workbook)
    }

    void test_read_empty_file()
    {
        std::string null_file = os.path.join(DATADIR, "reader", "null_file.xlsx");
        xlnt::workbook wb;
        TS_ASSERT_THROWS(InvalidFile, wb.load(null_file));
    }

        @raises(InvalidFileException)
    void test_read_empty_archive()
    {
        null_file = os.path.join(DATADIR, "reader", "null_archive.xlsx")
            wb = load_workbook(null_file)
    }

    void test_read_dimension()
    {
        path = os.path.join(DATADIR, "reader", "sheet2.xml")

            dimension = None
            handle = open(path)
            try :
            dimension = read_dimension(xml_source = handle.read())
            finally :
            handle.close()

            TS_ASSERT_EQUALS(("D", 1, "K", 30), dimension)
    }

    void test_calculate_dimension_iter()
    {
        path = os.path.join(DATADIR, "genuine", "empty.xlsx")
            wb = load_workbook(filename = path, use_iterators = True)
            sheet2 = wb.get_sheet_by_name("Sheet2 - Numbers")
            dimensions = sheet2.calculate_dimension()
            TS_ASSERT_EQUALS("%s%s:%s%s" % ("D", 1, "K", 30), dimensions)
    }

    void test_get_highest_row_iter()
    {
        path = os.path.join(DATADIR, "genuine", "empty.xlsx")
            wb = load_workbook(filename = path, use_iterators = True)
            sheet2 = wb.get_sheet_by_name("Sheet2 - Numbers")
            max_row = sheet2.get_highest_row()
            TS_ASSERT_EQUALS(30, max_row)
    }

    void test_read_workbook_with_no_properties()
    {
        genuine_wb = os.path.join(DATADIR, "genuine", \
            "empty_with_no_properties.xlsx")
            wb = load_workbook(filename = genuine_wb)
    }

    void setup_class_with_styles(cls)
    {
        cls.genuine_wb = os.path.join(DATADIR, "genuine", \
            "empty-with-styles.xlsx")
            wb = load_workbook(cls.genuine_wb)
            cls.ws = wb.get_sheet_by_name("Sheet1")
    }

    void test_read_general_style()
    {
        TS_ASSERT_EQUALS(ws.cell("A1").style.number_format.format_code,
            NumberFormat.FORMAT_GENERAL)
    }

    void test_read_date_style()
    {
        TS_ASSERT_EQUALS(ws.cell("A2").style.number_format.format_code,
            NumberFormat.FORMAT_DATE_XLSX14)
    }

    void test_read_number_style()
    {
        TS_ASSERT_EQUALS(ws.cell("A3").style.number_format.format_code,
            NumberFormat.FORMAT_NUMBER_00)
    }

    void test_read_time_style()
    {
        TS_ASSERT_EQUALS(ws.cell("A4").style.number_format.format_code,
            NumberFormat.FORMAT_DATE_TIME3)
    }

    void test_read_percentage_style()
    {
        TS_ASSERT_EQUALS(ws.cell("A5").style.number_format.format_code,
            NumberFormat.FORMAT_PERCENTAGE_00)
    }

    void setup_class_base_date_format(cls)
    {
        mac_wb_path = os.path.join(DATADIR, "reader", "date_1904.xlsx")
            cls.mac_wb = load_workbook(mac_wb_path)
            cls.mac_ws = cls.mac_wb.get_sheet_by_name("Sheet1")

            win_wb_path = os.path.join(DATADIR, "reader", "date_1900.xlsx")
            cls.win_wb = load_workbook(win_wb_path)
            cls.win_ws = cls.win_wb.get_sheet_by_name("Sheet1")
    }

    void test_read_win_base_date()
    {
        TS_ASSERT_EQUALS(win_wb.properties.excel_base_date, CALENDAR_WINDOWS_1900)
    }

    void test_read_mac_base_date()
    {
        TS_ASSERT_EQUALS(mac_wb.properties.excel_base_date, CALENDAR_MAC_1904)
    }

    void test_read_date_style_mac()
    {
        TS_ASSERT_EQUALS(mac_ws.cell("A1").style.number_format.format_code,
            NumberFormat.FORMAT_DATE_XLSX14)
    }

    void test_read_date_style_win()
    {
        TS_ASSERT_EQUALS(win_ws.cell("A1").style.number_format.format_code,
            NumberFormat.FORMAT_DATE_XLSX14)
    }

    void test_read_date_value()
    {
        datetuple = (2011, 10, 31)
            dt = datetime(datetuple[0], datetuple[1], datetuple[2])
            TS_ASSERT_EQUALS(mac_ws.cell("A1").value, dt)
            TS_ASSERT_EQUALS(win_ws.cell("A1").value, dt)
            TS_ASSERT_EQUALS(mac_ws.cell("A1").value, win_ws.cell("A1").value)
    }
};
