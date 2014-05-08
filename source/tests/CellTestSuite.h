#pragma once

#include <ctime>
#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/cell.h>
#include <xlnt/style.h>
#include <xlnt/workbook.h>
#include <xlnt/worksheet.h>

class CellTestSuite : public CxxTest::TestSuite
{
public:
    CellTestSuite()
    {

    }

    void test_coordinates()
    {
        auto coord = xlnt::cell::coordinate_from_string("ZF46");
        TS_ASSERT_EQUALS("ZF", coord.column);
        TS_ASSERT_EQUALS(46, coord.row);
    }

    void test_invalid_coordinate()
    {
        TS_ASSERT_THROWS_ANYTHING(xlnt::cell::coordinate_from_string("AAA"));
    }

    void test_zero_row()
    {
        TS_ASSERT_THROWS_ANYTHING(xlnt::cell::coordinate_from_string("AQ0"));
    }

    void test_absolute()
    {
        TS_ASSERT_EQUALS("$ZF$51", xlnt::cell::absolute_coordinate("ZF51"));
    }

    void test_absolute_multiple()
    {
        TS_ASSERT_EQUALS("$ZF$51:$ZF$53", xlnt::cell::absolute_coordinate("ZF51:ZF$53"));
    }

    void test_column_index()
    {
        TS_ASSERT_EQUALS(10, xlnt::cell::column_index_from_string("J"));
        TS_ASSERT_EQUALS(270, xlnt::cell::column_index_from_string("jJ"));
        TS_ASSERT_EQUALS(7030, xlnt::cell::column_index_from_string("jjj"));
    }


    void test_bad_column_index()
    {
        auto _check = [](const std::string &bad_string)
        {
            xlnt::cell::column_index_from_string(bad_string);
        };

        auto bad_strings = {"JJJJ", "", "$", "1"};
        for(auto bad_string : bad_strings)
        {
            TS_ASSERT_THROWS_ANYTHING(_check(bad_string));
        }
    }

    void test_column_letter_boundries()
    {
        TS_ASSERT_THROWS_ANYTHING(xlnt::cell::get_column_letter(0));
        TS_ASSERT_THROWS_ANYTHING(xlnt::cell::get_column_letter(18279));
    }


    void test_column_letter()
    {
        TS_ASSERT_EQUALS("ZZZ", xlnt::cell::get_column_letter(18278));
        TS_ASSERT_EQUALS("JJJ", xlnt::cell::get_column_letter(7030));
        TS_ASSERT_EQUALS("AB", xlnt::cell::get_column_letter(28));
        TS_ASSERT_EQUALS("AA", xlnt::cell::get_column_letter(27));
        TS_ASSERT_EQUALS("Z", xlnt::cell::get_column_letter(26));
    }


    void test_initial_value()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        xlnt::cell cell(ws, "A", 1, "17.5");

        TS_ASSERT_EQUALS(xlnt::cell::type::numeric, cell.get_data_type());
    }

    void test_null()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        xlnt::cell cell(ws, "A", 1);

        TS_ASSERT_EQUALS(xlnt::cell::type::null, cell.get_data_type());
    }

    void test_numeric()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        xlnt::cell cell(ws, "A", 1, "17.5");

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
        xlnt::worksheet ws(wb);
        xlnt::cell cell(ws, "A", 1);
        cell = "hello";
        TS_ASSERT_EQUALS(xlnt::cell::type::string, cell.get_data_type());
    }

    void test_single_dot()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        xlnt::cell cell(ws, "A", 1);
        cell = ".";
        TS_ASSERT_EQUALS(xlnt::cell::type::string, cell.get_data_type());
    }

    void test_formula()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        xlnt::cell cell(ws, "A", 1);
        cell = "=42";
        TS_ASSERT_EQUALS(xlnt::cell::type::formula, cell.get_data_type());
    }

    void test_boolean()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        xlnt::cell cell(ws, "A", 1);
        cell = true;
        TS_ASSERT_EQUALS(xlnt::cell::type::boolean, cell.get_data_type());
        cell = false;
        TS_ASSERT_EQUALS(xlnt::cell::type::boolean, cell.get_data_type());
    }

    void test_leading_zero()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        xlnt::cell cell(ws, "A", 1);
        cell = "0800";
        TS_ASSERT_EQUALS(xlnt::cell::type::string, cell.get_data_type());
    }

    void check_error()
    {
        //TS_ASSERT_EQUALS(xlnt::cell::type::error, cell.get_data_type());

        //for(error_string in cell.ERROR_CODES.keys())
        //{
        //    cell.value = error_string;
        //    yield check_error;
        //}
    }


    void test_data_type_check()
    {
        //xlnt::workbook wb;
        //xlnt::worksheet ws(wb);
        //xlnt::cell cell(ws, "A", 1);

        //cell.bind_value(xlnt::cell::$);
        //TS_ASSERT_EQUALS(xlnt::cell::type::null, cell.get_data_type());

        //cell.bind_value(".0e000");
        //TS_ASSERT_EQUALS(xlnt::cell::type::numeric, cell.get_data_type());

        //cell.bind_value("-0.e-0");
        //TS_ASSERT_EQUALS(xlnt::cell::type::numeric, cell.get_data_type());

        //cell.bind_value("1E");
        //TS_ASSERT_EQUALS(xlnt::cell::type::string, cell.get_data_type());
    }

    void test_set_bad_type()
    {
        //xlnt::workbook wb;
        //xlnt::worksheet ws(wb);
        //xlnt::cell cell(ws, "A", 1);

        //cell.set_value_explicit(1, "q");
    }


    void test_time()
    {
        //auto check_time = [](raw_value, coerced_value)
        //{
        //    cell.value = raw_value
        //        TS_ASSERT_EQUALS(cell.value, coerced_value)
        //        TS_ASSERT_EQUALS(cell.TYPE_NUMERIC, cell.data_type)
        //};

        //xlnt::workbook wb;
        //xlnt::worksheet ws(wb);
        //xlnt::cell cell(ws, "A", 1);

        //cell = "03:40:16";
        //TS_ASSERT_EQUALS(xlnt::cell::type::numeric, cell.get_data_type());
        //tm expected;
        //expected.tm_hour = 3;
        //expected.tm_min = 40;
        //expected.tm_sec = 16;
        //TS_ASSERT_EQUALS(cell, expected);

        //cell = "03:40";
        //TS_ASSERT_EQUALS(xlnt::cell::type::numeric, cell.get_data_type());
        //tm expected;
        //expected.tm_hour = 3;
        //expected.tm_min = 40;
        //TS_ASSERT_EQUALS(cell, expected);
    }

    void test_date_format_on_non_date()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        xlnt::cell cell(ws, "A", 1);
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
        xlnt::worksheet ws(wb);
        xlnt::cell cell(ws, "A", 1);
        cell = today;
        TS_ASSERT(today == cell);
    }

    void test_repr()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        xlnt::cell cell(ws, "A", 1);
        TS_ASSERT_EQUALS(cell.to_string(), "<Cell Sheet1.A1>");
    }

    void test_is_date()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        xlnt::cell cell(ws, "A", 1);

        time_t t = time(0);
        tm now = *localtime(&t);
        cell = now;
        TS_ASSERT_EQUALS(cell.is_date(), true);

        cell = "testme";
        TS_ASSERT("testme", cell);
        TS_ASSERT_EQUALS(cell.is_date(), false);
    }


    void test_is_not_date_color_format()
    {
        //xlnt::workbook wb;
        //xlnt::worksheet ws(wb);
        //xlnt::cell cell(ws, "A", 1);

        //cell = -13.5;
        //cell.get_style().get_number_format().set_format_code("0.00_);[Red]\(0.00\)");

        //TS_ASSERT_EQUALS(cell.is_date(), false);
    }
};
