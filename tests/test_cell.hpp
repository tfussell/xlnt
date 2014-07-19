#pragma once

#include <ctime>
#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/xlnt.hpp>

class test_cell : public CxxTest::TestSuite
{
public:
    void test_coordinates()
    {
        xlnt::cell_reference coord("ZF46");
        TS_ASSERT_EQUALS("ZF", coord.get_column());
        TS_ASSERT_EQUALS(46, coord.get_row());
    }

    void test_invalid_coordinate()
    {
        TS_ASSERT_THROWS(xlnt::cell_reference("AAA"), 
            xlnt::cell_coordinates_exception);
    }

    void test_zero_row()
    {
        TS_ASSERT_THROWS(xlnt::cell_reference("AQ0"), 
            xlnt::cell_coordinates_exception);
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
            {1, "A"},
            {10, "J"},
            {26, "Z"},
            {27, "AA"},
            {52, "AZ"},
            {53, "BA"},
            {78, "BZ"},
            {270, "jJ"},
            {677, "ZA"},
            {702, "ZZ"},
            {703, "AAA"},
            {728, "AAZ"},
            {731, "ABC"},
            {1353, "AZA"},
            {7030, "jjj"},
            {18253, "ZZA"},
            {18278, "ZZZ"}
        };

        for(auto expected_pair : expected)
        {
            TS_ASSERT_EQUALS(expected_pair.first, 
	        xlnt::cell_reference::column_index_from_string(expected_pair.second));
        }
    }

    void test_bad_column_index()
    {
        for(auto bad_string : {"JJJJ", "", "$", "1"})
        {
            TS_ASSERT_THROWS(xlnt::cell_reference::column_index_from_string(bad_string), xlnt::column_string_index_exception);
        }
    }

    void test_column_letter_boundaries()
    {
        TS_ASSERT_THROWS(xlnt::cell_reference::column_string_from_index(0),
	    xlnt::column_string_index_exception);
        TS_ASSERT_THROWS(xlnt::cell_reference::column_string_from_index(18279),
	    xlnt::column_string_index_exception);
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
        xlnt::worksheet ws = wb.create_sheet();
        xlnt::cell cell(ws, "A1", "17.5");

        TS_ASSERT_EQUALS(xlnt::cell::type::numeric, cell.get_data_type());
    }

    void test_1st()
    {
        xlnt::worksheet ws = wb.create_sheet();
        xlnt::cell cell(ws, "A1");

        TS_ASSERT_EQUALS(xlnt::cell::type::null, cell.get_data_type());
    }

    void test_null()
    {
        xlnt::worksheet ws = wb.create_sheet();
        xlnt::cell cell(ws, "A1", "17.5");
        cell.set_null();

        TS_ASSERT_EQUALS(xlnt::cell::type::null, cell.get_data_type());
    }

    void test_numeric()
    {
        xlnt::worksheet ws = wb.create_sheet();
        xlnt::cell cell(ws, "A1");

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
        xlnt::worksheet ws = wb.create_sheet();
        xlnt::cell cell(ws, "A1");

        cell = "hello";
        TS_ASSERT_EQUALS(xlnt::cell::type::string, cell.get_data_type());
    }

    void test_single_dot()
    {
        xlnt::worksheet ws = wb.create_sheet();
        xlnt::cell cell(ws, "A1");
        cell = ".";
        TS_ASSERT_EQUALS(xlnt::cell::type::string, cell.get_data_type());
    }

    void test_formula()
    {
        xlnt::worksheet ws = wb.create_sheet();
        xlnt::cell cell(ws, "A1");
        cell = "=42";
        TS_ASSERT_EQUALS(xlnt::cell::type::formula, cell.get_data_type());
        cell = "=if(A1<4;-1;1)";
        TS_ASSERT_EQUALS(xlnt::cell::type::formula, cell.get_data_type());
    }

    void test_boolean()
    {
        xlnt::worksheet ws = wb.create_sheet();
        xlnt::cell cell(ws, "A1");
        cell = true;
        TS_ASSERT_EQUALS(xlnt::cell::type::boolean, cell.get_data_type());
        cell = false;
        TS_ASSERT_EQUALS(xlnt::cell::type::boolean, cell.get_data_type());
    }

    void test_leading_zero()
    {
        xlnt::worksheet ws = wb.create_sheet();
        xlnt::cell cell(ws, "A1");
        cell = "0800";
        TS_ASSERT_EQUALS(xlnt::cell::type::string, cell.get_data_type());
    }

    void test_error_codes()
    {
        xlnt::worksheet ws = wb.create_sheet();
        xlnt::cell cell(ws, "A1");

        for(auto error : xlnt::cell::ErrorCodes)
        {
            cell = error.first;
            TS_ASSERT_EQUALS(xlnt::cell::type::error, cell.get_data_type());
        }
    }

    void test_insert_float()
    {
        xlnt::worksheet ws = wb.create_sheet();
        xlnt::cell cell(ws, "A1");
        cell = 3.14;
        TS_ASSERT_EQUALS(xlnt::cell::type::numeric, cell.get_data_type());
    }

    void test_insert_percentage()
    {
        xlnt::worksheet ws = wb.create_sheet();
        xlnt::cell cell(ws, "A1");
        cell = "3.14%";
        TS_ASSERT_DELTA(0.0314, cell.get_internal_value_numeric(), 1e-7);
    }

    void test_insert_datetime()
    {
        xlnt::worksheet ws = wb.create_sheet();
        xlnt::cell cell(ws, "A1");
        cell = xlnt::date::today();
        TS_ASSERT_EQUALS(xlnt::cell::type::numeric, cell.get_data_type());
    }

    void test_insert_date()
    {
        xlnt::worksheet ws = wb.create_sheet();
        xlnt::cell cell(ws, "A1");
        cell = xlnt::datetime::now();
        TS_ASSERT_EQUALS(xlnt::cell::type::numeric, cell.get_data_type());
    }

    void test_internal_date()
    {
        xlnt::worksheet ws = wb.create_sheet();
        xlnt::cell cell(ws, "A1");
        xlnt::datetime dt(2010, 7, 13, 6, 37, 41);
        cell = dt;
        TS_ASSERT_EQUALS(40372.27616898148, cell.get_internal_value_numeric());
    }

    void test_datetime_interpretation()
    {
        xlnt::worksheet ws = wb.create_sheet();
        xlnt::cell cell(ws, "A1");
        xlnt::datetime dt(2010, 7, 13, 6, 37, 41);
        cell = dt;
        TS_ASSERT_EQUALS(cell, dt);
        TS_ASSERT_DELTA(cell.get_internal_value_numeric(), 40372.27616898148, 1e-7);
    }

    void test_date_interpretation()
    {
        xlnt::worksheet ws = wb.create_sheet();
        xlnt::cell cell(ws, "A1");
        xlnt::date dt(2010, 7, 13);
        cell = dt;
        TS_ASSERT_EQUALS(cell, dt);
        TS_ASSERT_EQUALS(cell.get_internal_value_numeric(), 40372);
    }

    void test_number_format_style()
    {
        xlnt::worksheet ws = wb.create_sheet();
        xlnt::cell cell(ws, "A1");
        cell = "12.6%";
        TS_ASSERT_EQUALS(xlnt::number_format::format::percentage, cell.get_style().get_number_format().get_format_code());
    }

    void test_data_type_check()
    {
        xlnt::worksheet ws = wb.create_sheet();
        xlnt::cell cell(ws, "A1");

        TS_ASSERT_EQUALS(xlnt::cell::type::null, cell.get_data_type());

        cell = ".0e000";
        TS_ASSERT_EQUALS(xlnt::cell::type::numeric, cell.get_data_type());

        cell = "-0.e-0";
        TS_ASSERT_EQUALS(xlnt::cell::type::numeric, cell.get_data_type());

        cell = "1E";
        TS_ASSERT_EQUALS(xlnt::cell::type::string, cell.get_data_type());
    }

    void test_set_bad_type()
    {
        xlnt::worksheet ws = wb.create_sheet();
        xlnt::cell cell(ws, "A1");

        TS_ASSERT_THROWS(cell.set_explicit_value("1", xlnt::cell::type::formula),
	    xlnt::data_type_exception);
        TS_ASSERT_THROWS(cell.set_explicit_value(1, xlnt::cell::type::formula),
	    xlnt::data_type_exception);
        TS_ASSERT_THROWS(cell.set_explicit_value(1.0, xlnt::cell::type::formula),
	    xlnt::data_type_exception);
        TS_ASSERT_THROWS(cell.set_formula("1"), xlnt::data_type_exception);
        TS_ASSERT_THROWS(cell.set_error("1"), xlnt::data_type_exception);
        TS_ASSERT_THROWS(cell.set_hyperlink("1"), xlnt::data_type_exception);
        TS_ASSERT_THROWS(cell.set_formula("#REF!"), xlnt::data_type_exception);
    }


    void test_time()
    {
        xlnt::worksheet ws = wb.create_sheet();
        xlnt::cell cell(ws, "A1");

        cell = "03:40:16";
        TS_ASSERT_EQUALS(xlnt::cell::type::numeric, cell.get_data_type());
        TS_ASSERT_EQUALS(cell, xlnt::time(3, 40, 16));

        cell = "03:40";
        TS_ASSERT_EQUALS(xlnt::cell::type::numeric, cell.get_data_type());
        TS_ASSERT_EQUALS(cell, xlnt::time(3, 40));
    }
    
    void test_timedelta()
    {
        xlnt::worksheet ws = wb.create_sheet();
        xlnt::cell cell(ws, "A1");
        
        cell = xlnt::timedelta(1, 3, 0, 0, 0);
        TS_ASSERT_EQUALS(cell, 1.125);
        TS_ASSERT_EQUALS(cell.get_data_type(), xlnt::cell::type::numeric);
    }

    void test_date_format_on_non_date()
    {
        xlnt::worksheet ws = wb.create_sheet();
        xlnt::cell cell(ws, "A1");

	cell = xlnt::datetime::now();
        cell = "testme";
        TS_ASSERT("testme" == cell);
    }

    void test_set_get_date()
    {
        xlnt::datetime today(2010, 1, 18, 14,15, 20 ,1600);

        xlnt::worksheet ws = wb.create_sheet();
        xlnt::cell cell(ws, "A1");

        cell = today;
        TS_ASSERT(today == cell);
    }

    void test_repr()
    {
        xlnt::worksheet ws = wb.create_sheet();
        xlnt::cell cell(ws, "A1");

        TS_ASSERT_EQUALS(cell.to_string(), "<Cell " + ws.get_title() + ".A1>");
    }

    void test_is_date()
    {
        xlnt::worksheet ws = wb.create_sheet();
        xlnt::cell cell(ws, "A1");

        cell = xlnt::datetime::now();
        TS_ASSERT(cell.is_date());

        cell = "testme";
        TS_ASSERT_EQUALS("testme", cell);
        TS_ASSERT(!cell.is_date());
    }


    void test_is_not_date_color_format()
    {
        xlnt::worksheet ws = wb.create_sheet();
        xlnt::cell cell(ws, "A1");

        cell = -13.5;
        cell.get_style().get_number_format().set_format_code("0.00_);[Red]\\(0.00\\)");

        TS_ASSERT(!cell.is_date());
    }
    
    void test_comment_count()
    {
        xlnt::worksheet ws = wb.create_sheet();
        xlnt::cell cell(ws, "A1");
        
        TS_ASSERT(ws.get_comment_count() == 0);
        cell.set_comment(xlnt::comment("text", "author"));
        TS_ASSERT(ws.get_comment_count() == 1);
        cell.set_comment(xlnt::comment("text", "author"));
        TS_ASSERT(ws.get_comment_count() == 1);
        cell.clear_comment();
        TS_ASSERT(ws.get_comment_count() == 0);
        cell.clear_comment();
        TS_ASSERT(ws.get_comment_count() == 0);
    }
    
    void test_comment_assignment()
    {
        xlnt::worksheet ws = wb.create_sheet();
        xlnt::cell cell(ws, "A1");
        
        xlnt::comment c("text", "author");
        cell.set_comment(c);
        TS_ASSERT_THROWS_ANYTHING(ws.get_cell("A2").set_comment(c));
        ws.get_cell("A2").set_comment(xlnt::comment("text2", "author2"));
        TS_ASSERT_THROWS_ANYTHING(ws.get_cell("A1").set_comment(ws.get_cell("A2").get_comment()));
        ws.get_cell("A1").clear_comment();
        TS_ASSERT_THROWS_NOTHING(ws.get_cell("A2").set_comment(c));
    }
    
    void test_cell_offset()
    {
        xlnt::worksheet ws = wb.create_sheet();
        TS_ASSERT(ws.get_cell("B15").offset(1, 2).get_reference() == "C17");
    }

private:
    xlnt::workbook wb;
};
