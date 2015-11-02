#pragma once

#include <ctime>
#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/cell/cell.hpp>
#include <xlnt/cell/cell_reference.hpp>
#include <xlnt/cell/comment.hpp>
#include <xlnt/common/datetime.hpp>
#include <xlnt/common/exceptions.hpp>
#include <xlnt/styles/alignment.hpp>
#include <xlnt/styles/border.hpp>
#include <xlnt/styles/font.hpp>
#include <xlnt/styles/fill.hpp>
#include <xlnt/styles/number_format.hpp>
#include <xlnt/styles/protection.hpp>
#include <xlnt/worksheet/range.hpp>
#include <xlnt/worksheet/worksheet.hpp>
#include <xlnt/workbook/workbook.hpp>

class test_cell : public CxxTest::TestSuite
{
private:
    xlnt::workbook wb, wb_guess_types;
    
public:
    test_cell()
    {
        wb_guess_types.set_guess_types(true);
    }
    
	void test_infer_numeric()
	{
		auto ws = wb_guess_types.create_sheet();
        auto cell = ws.get_cell("A1");

		cell.set_value("4.2");
		TS_ASSERT(cell.get_value<long double>() == 4.2L);

		cell.set_value("-42.000");
		TS_ASSERT(cell.get_value<int>() == -42);

		cell.set_value("0");
		TS_ASSERT(cell.get_value<int>() == 0);

		cell.set_value("0.9999");
		TS_ASSERT(cell.get_value<long double>() == 0.9999L);

		cell.set_value("99E-02");
		TS_ASSERT(cell.get_value<long double>() == 0.99L);

		cell.set_value("4");
		TS_ASSERT(cell.get_value<int>() == 4);

		cell.set_value("-1E3");
		TS_ASSERT(cell.get_value<int>() == -1000);

		cell.set_value("2e+2");
		TS_ASSERT(cell.get_value<int>() == 200);

		cell.set_value("3.1%");
		TS_ASSERT(cell.get_value<long double>() == 0.031L);

		cell.set_value("03:40:16");
        TS_ASSERT(cell.get_value<xlnt::time>() == xlnt::time(3, 40, 16));

		cell.set_value("03:40");
        TS_ASSERT(cell.get_value<xlnt::time>() == xlnt::time(3, 40));

		cell.set_value("30:33.865633336");
        TS_ASSERT(cell.get_value<xlnt::time>() == xlnt::time(0, 30, 33, 865633));
	}

    void test_ctor()
    {
        auto ws = wb.create_sheet();
        auto cell = ws.get_cell(xlnt::cell_reference("A", 1));
        
        TS_ASSERT(cell.get_data_type() == xlnt::cell::type::null);
        TS_ASSERT(cell.get_column() == "A");
        TS_ASSERT(cell.get_row() == 1);
        TS_ASSERT(cell.get_reference() == "A1");
        TS_ASSERT(!cell.has_value());
        TS_ASSERT(!cell.has_comment());
    }

    
    void test_null()
    {
        auto datatypes =
        {
            xlnt::cell::type::null
        };
        
        for(auto datatype : datatypes)
        {
            auto ws = wb.create_sheet();
            auto cell = ws.get_cell(xlnt::cell_reference(1, 1));
            
            cell.set_data_type(datatype);
            TS_ASSERT(cell.get_data_type() == datatype);
            cell.clear_value();
            TS_ASSERT(cell.get_data_type() == xlnt::cell::type::null);
        }
    }
    
    void test_string()
    {
        auto ws = wb.create_sheet();
        auto cell = ws.get_cell(xlnt::cell_reference(1, 1));
        
        cell.set_value("hello");
        TS_ASSERT(cell.get_data_type() == xlnt::cell::type::string);
        
        cell.set_value(".");
        TS_ASSERT(cell.get_data_type() == xlnt::cell::type::string);
        
        cell.set_value("0800");
        TS_ASSERT(cell.get_data_type() == xlnt::cell::type::string);
    }
    
    void test_formula1()
    {
        auto ws = wb_guess_types.create_sheet();
        auto cell = ws.get_cell(xlnt::cell_reference(1, 1));
        
        cell.set_value("=42");
        TS_ASSERT(cell.get_data_type() == xlnt::cell::type::formula);
    }
    
    void test_formula2()
    {
        auto ws = wb.create_sheet();
        auto cell = ws.get_cell(xlnt::cell_reference(1, 1));
        
        cell.set_value("=if(A1<4;-1;1)");
        TS_ASSERT(cell.get_data_type() == xlnt::cell::type::formula);
    }
    
    void test_not_formula()
    {
        auto ws = wb.create_sheet();
        auto cell = ws.get_cell(xlnt::cell_reference(1, 1));
        
        cell.set_value("=");
        TS_ASSERT(cell.get_data_type() == xlnt::cell::type::string);
        TS_ASSERT(cell.get_value<std::string>() == "=");
        TS_ASSERT(!cell.has_formula());
    }
    
    void test_boolean()
    {
        auto ws = wb.create_sheet();
        auto cell = ws.get_cell(xlnt::cell_reference(1, 1));
        
        for(auto value : {true, false})
        {
            cell.set_value(value);
            TS_ASSERT(cell.get_data_type() == xlnt::cell::type::boolean);
        }
    }
    
    void test_error_codes()
    {
        auto ws = wb_guess_types.create_sheet();
        auto cell = ws.get_cell(xlnt::cell_reference(1, 1));
        
        for(auto error_code : xlnt::cell::error_codes())
        {
            cell.set_value(error_code.first);
            TS_ASSERT(cell.get_data_type() == xlnt::cell::type::error);
        }
    }

    void test_insert_datetime()
    {
        auto ws = wb.create_sheet();
        auto cell = ws.get_cell(xlnt::cell_reference(1, 1));
        
        cell.set_value(xlnt::datetime(2010, 7, 13, 6, 37, 41));
        TS_ASSERT(cell.get_data_type() == xlnt::cell::type::numeric);
        TS_ASSERT(cell.get_value<long double>() == 40372.27616898148L);
        TS_ASSERT(cell.is_date());
        TS_ASSERT(cell.get_number_format().get_format_string() == "yyyy-mm-dd h:mm:ss");
    }
    
    void test_insert_date()
    {
        auto ws = wb.create_sheet();
        auto cell = ws.get_cell(xlnt::cell_reference(1, 1));
        
        cell.set_value(xlnt::date(2010, 7, 13));
        TS_ASSERT(cell.get_data_type() == xlnt::cell::type::numeric);
        TS_ASSERT(cell.get_value<long double>() == 40372.L);
        TS_ASSERT(cell.is_date());
        TS_ASSERT(cell.get_number_format().get_format_string() == "yyyy-mm-dd");
    }
    
    void test_insert_time()
    {
        auto ws = wb.create_sheet();
        auto cell = ws.get_cell(xlnt::cell_reference(1, 1));
        
        cell.set_value(xlnt::time(1, 3));
        TS_ASSERT(cell.get_data_type() == xlnt::cell::type::numeric);
        TS_ASSERT(cell.get_value<long double>() == 0.04375L);
        TS_ASSERT(cell.is_date());
        TS_ASSERT(cell.get_number_format().get_format_string() == "h:mm:ss");
    }
    
    void test_cell_formatted_as_date1()
    {
        auto ws = wb.create_sheet();
        auto cell = ws.get_cell(xlnt::cell_reference(1, 1));
        
        cell.set_value(xlnt::datetime::today());
        cell.clear_value();
        TS_ASSERT(!cell.is_date()); // disagree with openpyxl
        TS_ASSERT(!cell.has_value());
    }
    
    void test_cell_formatted_as_date2()
    {
        auto ws = wb.create_sheet();
        auto cell = ws.get_cell(xlnt::cell_reference(1, 1));
        
        cell.set_value(xlnt::datetime::today());
        cell.set_value("testme");
        TS_ASSERT(!cell.is_date());
        TS_ASSERT(cell.get_value<std::string>() == "testme");
    }
    
    void test_cell_formatted_as_date3()
    {
        auto ws = wb.create_sheet();
        auto cell = ws.get_cell(xlnt::cell_reference(1, 1));
        
        cell.set_value(xlnt::datetime::today());
        cell.set_value(true);
        TS_ASSERT(!cell.is_date());
        TS_ASSERT(cell.get_value<bool>() == true);
    }
    
    void test_illegal_chacters()
    {
        auto ws = wb.create_sheet();
        auto cell = ws.get_cell(xlnt::cell_reference(1, 1));
    
        // The bytes 0x00 through 0x1F inclusive must be manually escaped in values.
        auto illegal_chrs = {0, 1, 2, 3, 4, 5, 6, 7, 8, 11, 12, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31};
        
        for(auto i : illegal_chrs)
        {
            std::string str(1, i);
            TS_ASSERT_THROWS(cell.set_value(str), xlnt::illegal_character_error);
        }
        
        cell.set_value(std::string(1, 33));
        cell.set_value(std::string(1, 9));  // Tab
        cell.set_value(std::string(1, 10));  // Newline
        cell.set_value(std::string(1, 13));  // Carriage return
        cell.set_value(" Leading and trailing spaces are legal ");
    }
    /*
    values = (
              ('30:33.865633336', [('', '', '', '30', '33', '865633')]),
              ('03:40:16', [('03', '40', '16', '', '', '')]),
              ('03:40', [('03', '40', '',  '', '', '')]),
              ('55:72:12', []),
              )
    @pytest.mark.parametrize("value, expected",
                             values)
     */
    void test_time_regex()
    {
        /*
        from openpyxl.cell.cell import TIME_REGEX;
        m = TIME_REGEX.findall(value);
        TS_ASSERT(m == expected;
         */
    }
    
    void test_timedelta()
    {
        auto ws = wb.create_sheet();
        auto cell = ws.get_cell(xlnt::cell_reference(1, 1));
        
        cell.set_value(xlnt::timedelta(1, 3));
        TS_ASSERT(cell.get_value<long double>() == 1.125);
        TS_ASSERT(cell.get_data_type() == xlnt::cell::type::numeric);
        TS_ASSERT(!cell.is_date());
        TS_ASSERT(cell.get_number_format().get_format_string() == "[hh]:mm:ss");
    }
    
    void test_repr()
    {
        auto ws = wb[1];
        auto cell = ws.get_cell(xlnt::cell_reference(1, 1));
        
        TS_ASSERT(cell.to_repr() == "<Cell Sheet1.A1>");
    }
    
    void test_comment_assignment()
    {
        auto ws = wb.create_sheet();
        auto cell = ws.get_cell(xlnt::cell_reference(1, 1));
        
        TS_ASSERT(!cell.has_comment());
        xlnt::comment comm(cell, "text", "author");
        TS_ASSERT(cell.get_comment() == comm);
    }
    
    void test_comment_count()
    {
        auto ws = wb.create_sheet();
        auto cell = ws.get_cell(xlnt::cell_reference(1, 1));

        TS_ASSERT(ws.get_comment_count() == 0);
        xlnt::comment(cell, "text", "author");
        TS_ASSERT(ws.get_comment_count() == 1);
        xlnt::comment(cell, "text", "author");
        TS_ASSERT(ws.get_comment_count() == 1);
        cell.clear_comment();
        TS_ASSERT(ws.get_comment_count() == 0);
        cell.clear_comment();
        TS_ASSERT(ws.get_comment_count() == 0);
    }
    
    void test_only_one_cell_per_comment()
    {
        auto ws = wb.create_sheet();
        auto cell = ws.get_cell(xlnt::cell_reference(1, 1));
        xlnt::comment comm(cell, "text", "author");
    
        auto c2 = ws.get_cell(xlnt::cell_reference(1, 2));
        TS_ASSERT_THROWS(c2.set_comment(comm), xlnt::attribute_error);
    }
    
    void test_remove_comment()
    {
        auto ws = wb.create_sheet();
        auto cell = ws.get_cell(xlnt::cell_reference(1, 1));
        
        xlnt::comment comm(cell, "text", "author");
        cell.clear_comment();
        TS_ASSERT(!cell.has_comment());
    }
    
    void test_cell_offset()
    {
        /*
        auto ws = wb.create_sheet();
        auto cell = ws.get_cell(xlnt::cell_reference(1, 1));
        TS_ASSERT(cell.offset(2, 1).get_reference() == "B3");
         */
    }
    
    void test_bad_encoding()
    {
        /*
        unsigned char pound = 163;
        auto test_string = "Compount Value (" + std::string(pound) + ")";
        auto ws = wb.create_sheet();
        cell = ws[xlnt::cell_reference("A1")];
        TS_ASSERT_THROWS(cell.check_string(test_string), xlnt::unicode_decode_error);
        TS_ASSERT_THROWS(cell.set_value(test_string), xlnt::unicode_decode_error);
         */
    }
    
    void test_good_encoding()
    {
        /*
        auto wb = xlnt::workbook(xlnt::encoding::latin1);
        auto ws = wb.get_active_sheet();
        auto cell = ws[xlnt::cell_reference("A1")];
        cell.set_value(test_string);
         */
    }
    
    void _test_font()
    {
        xlnt::font font;
        font.set_bold(true);
        auto ws = wb.create_sheet();
        ws.get_parent().add_font(font);
    
        auto cell = xlnt::cell(ws, "A1");
        TS_ASSERT_EQUALS(cell.get_font(), font);
    }
    
    void _test_fill()
    {
        xlnt::fill f;
        f.set_type(xlnt::fill::type::solid);
//        f.set_foreground_color("FF0000");
        auto ws = wb.create_sheet();
        ws.get_parent().add_fill(f);
    
        xlnt::cell cell(ws, "A1");
        TS_ASSERT(cell.get_fill() == f);
    }
    
    void _test_border()
    {
        xlnt::border border;
        auto ws = wb.create_sheet();
        ws.get_parent().add_border(border);
        
        auto cell = ws.get_cell(xlnt::cell_reference(1, 1));
        TS_ASSERT(cell.get_border() == border);
    }
    
    void _test_number_format()
    {
        auto ws = wb.create_sheet();
        ws.get_parent().add_number_format(xlnt::number_format("dd--hh--mm"));
    
        xlnt::cell cell(ws, "A1");
        cell.set_number_format(xlnt::number_format("dd--hh--mm"));
        TS_ASSERT(cell.get_number_format().get_format_string() == "dd--hh--mm");
    }
    
    void _test_alignment()
    {
        xlnt::alignment align;
        align.set_wrap_text(true);
        
        auto ws = wb.create_sheet();
        ws.get_parent().add_alignment(align);
    
        xlnt::cell cell(ws, "A1");
        TS_ASSERT(cell.get_alignment() == align);
    }
    
    void _test_protection()
    {
        xlnt::protection prot;
        prot.set_locked(xlnt::protection::type::protected_);
        
        auto ws = wb.create_sheet();
        ws.get_parent().add_protection(prot);
    
        xlnt::cell cell(ws, "A1");
        TS_ASSERT(cell.get_protection() == prot);
    }
    
    void _test_pivot_button()
    {
        auto ws = wb.create_sheet();
        
        xlnt::cell cell(ws, "A1");
        cell.set_pivot_button(true);
        TS_ASSERT(cell.pivot_button());
    }
    
    void _test_quote_prefix()
    {
        auto ws = wb.create_sheet();
    
        xlnt::cell cell(ws, "A1");
        cell.set_quote_prefix(true);
        TS_ASSERT(cell.quote_prefix());
    }
};
