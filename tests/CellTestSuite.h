#pragma once

#include <ctime>
#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/xlnt.h>

class CellTestSuite : public CxxTest::TestSuite
{
public:
    CellTestSuite()
    {
    }

    void test_coordinates()
    {
        xlnt::cell_reference coord("ZF46");
        TS_ASSERT_EQUALS("ZF", coord.get_column());
        TS_ASSERT_EQUALS(46, coord.get_row());
    }

    void test_invalid_coordinate()
    {
        TS_ASSERT_THROWS_ANYTHING(xlnt::cell_reference("AAA"));
    }

    void test_zero_row()
    {
        TS_ASSERT_THROWS_ANYTHING(xlnt::cell_reference("AQ0"));
    }

    void test_absolute()
    {
        TS_ASSERT_EQUALS("$ZF$51", xlnt::cell_reference::make_absolute("ZF51").to_string());
    }

    void test_absolute_multiple()
    {
        TS_ASSERT_EQUALS("$ZF$51:$ZF$53", xlnt::range_reference::make_absolute("ZF51:ZF$53").to_string());
    }

    void test_column_index()
    {
        static const std::unordered_map<int, std::string> expected = 
        {
            {10, "J"},
            {270, "jJ"},
            {7030, "jjj"},
            {1, "A"},
            {26, "Z"},
            {27, "AA"},
            {52, "AZ"},
            {53, "BA"},
            {78, "BZ"},
            {677, "ZA"},
            {702, "ZZ"},
            {703, "AAA"},
            {728, "AAZ"},
            {731, "ABC"},
            {1353, "AZA"},
            {18253, "ZZA"},
            {18278, "ZZZ"}
        };

        for(auto expected_pair : expected)
        {
            auto calculated = xlnt::cell_reference::column_index_from_string(expected_pair.second);
            TS_ASSERT_EQUALS(expected_pair.first, calculated);
        }
    }

    void test_bad_column_index()
    {
	for(auto bad_string : {"JJJJ", "", "$", "1"})
        {
            TS_ASSERT_THROWS_ANYTHING(xlnt::cell_reference::column_index_from_string(bad_string));
        }
    }

    void test_column_letter_boundries()
    {
        TS_ASSERT_THROWS_ANYTHING(xlnt::cell_reference::column_string_from_index(0));
        TS_ASSERT_THROWS_ANYTHING(xlnt::cell_reference::column_string_from_index(18279));
    }


    void test_column_letter()
    {
        TS_ASSERT_EQUALS("ZZZ", xlnt::cell_reference::column_string_from_index(18278));
        TS_ASSERT_EQUALS("JJJ", xlnt::cell_reference::column_string_from_index(7030));
        TS_ASSERT_EQUALS("AB", xlnt::cell_reference::column_string_from_index(28));
        TS_ASSERT_EQUALS("AA", xlnt::cell_reference::column_string_from_index(27));
        TS_ASSERT_EQUALS("Z", xlnt::cell_reference::column_string_from_index(26));
    }


    void test_initial_value()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws = wb.get_active_sheet();
        xlnt::cell cell(ws, "A1", "17.5");

        TS_ASSERT_EQUALS(xlnt::cell::type::numeric, cell.get_data_type());
    }

    void test_null()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws = wb.get_active_sheet();
        xlnt::cell cell(ws, "A1");

        TS_ASSERT_EQUALS(xlnt::cell::type::null, cell.get_data_type());
    }

    void test_numeric()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws = wb.get_active_sheet();
        xlnt::cell cell(ws, "A1", "17.5");

        cell = 42;
        TS_ASSERT_EQUALS(xlnt::cell::type::numeric, cell.get_data_type());
        cell = "4.2";
        TS_ASSERT_EQUALS(xlnt::cell::type::numeric, cell.get_data_type());
        cell = "-42.000";
        TS_ASSERT_EQUALS(xlnt::cell::type::numeric, cell.get_data_type());
        cell = "0";
        TS_ASSERT_EQUALS(xlnt::cell::type::numeric, cell.get_data_type());
        cell = 0;
        TS_ASSERT_EQUALS(xlnt::cell::type::numeric, cell.get_data_type());
        cell = 0.0001;
        TS_ASSERT_EQUALS(xlnt::cell::type::numeric, cell.get_data_type());
        cell = "0.9999";
        TS_ASSERT_EQUALS(xlnt::cell::type::numeric, cell.get_data_type());
        cell = "99E-02";
        TS_ASSERT_EQUALS(xlnt::cell::type::numeric, cell.get_data_type());
        cell = 1e1;
        TS_ASSERT_EQUALS(xlnt::cell::type::numeric, cell.get_data_type());
        cell = "4";
        TS_ASSERT_EQUALS(xlnt::cell::type::numeric, cell.get_data_type());
        cell = "-1E3";
        TS_ASSERT_EQUALS(xlnt::cell::type::numeric, cell.get_data_type());
        cell = 4;
        TS_ASSERT_EQUALS(xlnt::cell::type::numeric, cell.get_data_type());
    }

    void test_string()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws = wb.get_active_sheet();
        xlnt::cell cell(ws, "A1");

        cell = "hello";
        TS_ASSERT_EQUALS(xlnt::cell::type::string, cell.get_data_type());
    }

    void test_single_dot()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws = wb.get_active_sheet();
        xlnt::cell cell(ws, "A1");
        cell = ".";
        TS_ASSERT_EQUALS(xlnt::cell::type::string, cell.get_data_type());
    }

    void test_formula()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws = wb.get_active_sheet();
        xlnt::cell cell(ws, "A1");
        cell = "=42";
        TS_ASSERT_EQUALS(xlnt::cell::type::formula, cell.get_data_type());
    }

    void test_boolean()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws = wb.get_active_sheet();
        xlnt::cell cell(ws, "A1");
        cell = true;
        TS_ASSERT_EQUALS(xlnt::cell::type::boolean, cell.get_data_type());
        cell = false;
        TS_ASSERT_EQUALS(xlnt::cell::type::boolean, cell.get_data_type());
    }

    void test_leading_zero()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws = wb.get_active_sheet();
        xlnt::cell cell(ws, "A1");
        cell = "0800";
        TS_ASSERT_EQUALS(xlnt::cell::type::string, cell.get_data_type());
    }

    void test_error_codes()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws = wb.get_active_sheet();
        xlnt::cell cell(ws, "A1");

        for(auto error : xlnt::cell::ErrorCodes)
        {
            cell = error.first;
            TS_ASSERT_EQUALS(xlnt::cell::type::error, cell.get_data_type());
        }
    }


    void test_data_type_check()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws = wb.get_active_sheet();
        xlnt::cell cell(ws, "A1");

        TS_ASSERT_EQUALS(xlnt::cell::type::null, cell.get_data_type());

        cell.bind_value(".0e000");
        TS_ASSERT_EQUALS(xlnt::cell::type::numeric, cell.get_data_type());

        cell.bind_value("-0.e-0");
        TS_ASSERT_EQUALS(xlnt::cell::type::numeric, cell.get_data_type());

        cell.bind_value("1E");
        TS_ASSERT_EQUALS(xlnt::cell::type::string, cell.get_data_type());
    }

    void test_set_bad_type()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws = wb.get_active_sheet();
        xlnt::cell cell(ws, "A1");

        cell.set_explicit_value("1", xlnt::cell::type::formula);
    }


    void test_time()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws = wb.get_active_sheet();
        xlnt::cell cell(ws, "A1");

        cell = "03:40:16";
        TS_ASSERT_EQUALS(xlnt::cell::type::date, cell.get_data_type());
        tm expected1;
        expected1.tm_hour = 3;
        expected1.tm_min = 40;
        expected1.tm_sec = 16;
        TS_ASSERT(cell == expected1);

        cell = "03:40";
        TS_ASSERT_EQUALS(xlnt::cell::type::date, cell.get_data_type());
        tm expected2;
        expected2.tm_hour = 3;
        expected2.tm_min = 40;
        TS_ASSERT(cell == expected1);
    }

    void test_date_format_on_non_date()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws = wb.get_active_sheet();
        xlnt::cell cell(ws, "A1");

        time_t t = time(0);
        tm now = *localtime(&t);
        cell = now;
        cell = "testme";
        TS_ASSERT("testme" == cell);
    }

    void test_set_get_date()
    {
        tm today = {0};
        today.tm_year = 2010;
        today.tm_mon = 1;
        today.tm_yday = 18;
        today.tm_hour = 14;
        today.tm_min = 15;
        today.tm_sec = 20;

        xlnt::workbook wb;
        xlnt::worksheet ws = wb.get_active_sheet();
        xlnt::cell cell(ws, "A1");

        cell = today;
        TS_ASSERT(today == cell);
    }

    void test_repr()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws = wb.get_active_sheet();
        xlnt::cell cell(ws, "A1");

        TS_ASSERT_EQUALS(cell.to_string(), "<Cell A1>");
    }

    void test_is_date()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws = wb.get_active_sheet();
        xlnt::cell cell(ws, "A1");

        time_t t = time(0);
        tm now = *localtime(&t);
        cell = now;
        TS_ASSERT_EQUALS(cell.is_date(), true);

        cell = "testme";
        TS_ASSERT_EQUALS("testme", cell);
        TS_ASSERT_EQUALS(cell.is_date(), false);
    }


    void test_is_not_date_color_format()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws = wb.get_active_sheet();
        xlnt::cell cell(ws, "A1");

        cell = -13.5;
        //cell.get_style().get_number_format().set_format_code("0.00_);[Red]\\(0.00\\)");

        TS_ASSERT_EQUALS(cell.is_date(), false);
    }
};
