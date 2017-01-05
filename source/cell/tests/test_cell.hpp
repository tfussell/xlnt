#pragma once

#include <cmath>
#include <ctime>
#include <fstream>
#include <iostream>
#include <sstream>
#include <cxxtest/TestSuite.h>

#include <xlnt/xlnt.hpp>

class test_cell : public CxxTest::TestSuite
{
public:
	void test_infer_numeric()
	{
        xlnt::workbook wb;
		auto ws = wb.active_sheet();
        auto cell = ws.cell("A1");

		cell.value("4.2", true);
		TS_ASSERT(cell.value<long double>() == 4.2L);

		cell.value("-42.000", true);
		TS_ASSERT(cell.value<int>() == -42);

		cell.value("0", true);
		TS_ASSERT(cell.value<int>() == 0);

		cell.value("0.9999", true);
		TS_ASSERT(cell.value<long double>() == 0.9999L);

		cell.value("99E-02", true);
		TS_ASSERT(cell.value<long double>() == 0.99L);

		cell.value("4", true);
		TS_ASSERT(cell.value<int>() == 4);

		cell.value("-1E3", true);
		TS_ASSERT(cell.value<int>() == -1000);

		cell.value("2e+2", true);
		TS_ASSERT(cell.value<int>() == 200);

		cell.value("3.1%", true);
		TS_ASSERT(cell.value<long double>() == 0.031L);

		cell.value("03:40:16", true);
        TS_ASSERT(cell.value<xlnt::time>() == xlnt::time(3, 40, 16));

		cell.value("03:", true);
        TS_ASSERT_EQUALS(cell.value<std::string>(), "03:");

		cell.value("03:40", true);
        TS_ASSERT(cell.value<xlnt::time>() == xlnt::time(3, 40));

		cell.value("30:33.865633336", true);
        TS_ASSERT(cell.value<xlnt::time>() == xlnt::time(0, 30, 33, 865633));
	}

    void test_ctor()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        auto cell = ws.cell(xlnt::cell_reference("A", 1));
        
        TS_ASSERT(cell.data_type() == xlnt::cell::type::null);
        TS_ASSERT(cell.column() == "A");
        TS_ASSERT(cell.row() == 1);
        TS_ASSERT(cell.reference() == "A1");
        TS_ASSERT(!cell.has_value());
    }

    void test_null()
    {
        xlnt::workbook wb;

        const auto datatypes =
        {
            xlnt::cell::type::null,
            xlnt::cell::type::boolean,
            xlnt::cell::type::error,
            xlnt::cell::type::formula,
            xlnt::cell::type::numeric,
            xlnt::cell::type::string
        };
        
        for(const auto &datatype : datatypes)
        {
            auto ws = wb.active_sheet();
            auto cell = ws.cell(xlnt::cell_reference(1, 1));
            
            cell.data_type(datatype);
            TS_ASSERT(cell.data_type() == datatype);
            cell.clear_value();
            TS_ASSERT(cell.data_type() == xlnt::cell::type::null);
        }
    }
    
    void test_string()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        auto cell = ws.cell(xlnt::cell_reference(1, 1));
        
        cell.value("hello");
        TS_ASSERT(cell.data_type() == xlnt::cell::type::string);
        
        cell.value(".");
        TS_ASSERT(cell.data_type() == xlnt::cell::type::string);
        
        cell.value("0800");
        TS_ASSERT(cell.data_type() == xlnt::cell::type::string);
    }
    
    void test_formula1()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        auto cell = ws.cell(xlnt::cell_reference(1, 1));

        cell.value("=42", true);
        TS_ASSERT(cell.data_type() == xlnt::cell::type::formula);
    }
    
    void test_formula2()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        auto cell = ws.cell(xlnt::cell_reference(1, 1));
        
        cell.value("=if(A1<4;-1;1)", true);
        TS_ASSERT(cell.data_type() == xlnt::cell::type::formula);
    }

    void test_formula3()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        auto cell = ws.cell(xlnt::cell_reference(1, 1));

        TS_ASSERT(!cell.has_formula());
        TS_ASSERT_THROWS(cell.formula(""), xlnt::invalid_parameter);
        TS_ASSERT(!cell.has_formula());
        cell.formula("=42");
        TS_ASSERT(cell.has_formula());
        TS_ASSERT_EQUALS(cell.formula(), "42");
        cell.clear_formula();
        TS_ASSERT(!cell.has_formula());
    }

    
    void test_not_formula()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        auto cell = ws.cell(xlnt::cell_reference(1, 1));
        
        cell.value("=");
        TS_ASSERT(cell.data_type() == xlnt::cell::type::string);
        TS_ASSERT(cell.value<std::string>() == "=");
        TS_ASSERT(!cell.has_formula());
    }
    
    void test_boolean()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        auto cell = ws.cell(xlnt::cell_reference(1, 1));
        
        for(auto value : {true, false})
        {
            cell.value(value);
            TS_ASSERT(cell.data_type() == xlnt::cell::type::boolean);
        }
    }
    
    void test_error_codes()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        auto cell = ws.cell(xlnt::cell_reference(1, 1));
        
        for(auto error_code : xlnt::cell::error_codes())
        {
            cell.value(error_code.first, true);
            TS_ASSERT(cell.data_type() == xlnt::cell::type::error);
        }
    }

    void test_insert_datetime()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        auto cell = ws.cell(xlnt::cell_reference(1, 1));
        
        cell.value(xlnt::datetime(2010, 7, 13, 6, 37, 41));
        TS_ASSERT(cell.data_type() == xlnt::cell::type::numeric);
        TS_ASSERT(cell.value<long double>() == 40372.27616898148L);
        TS_ASSERT(cell.is_date());
        TS_ASSERT(cell.number_format().format_string() == "yyyy-mm-dd h:mm:ss");
    }
    
    void test_insert_date()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        auto cell = ws.cell(xlnt::cell_reference(1, 1));
        
        cell.value(xlnt::date(2010, 7, 13));
        TS_ASSERT(cell.data_type() == xlnt::cell::type::numeric);
        TS_ASSERT(cell.value<long double>() == 40372.L);
        TS_ASSERT(cell.is_date());
        TS_ASSERT(cell.number_format().format_string() == "yyyy-mm-dd");
    }
    
    void test_insert_time()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        auto cell = ws.cell(xlnt::cell_reference(1, 1));
        
        cell.value(xlnt::time(1, 3));
        TS_ASSERT(cell.data_type() == xlnt::cell::type::numeric);
        TS_ASSERT(cell.value<long double>() == 0.04375L);
        TS_ASSERT(cell.is_date());
        TS_ASSERT(cell.number_format().format_string() == "h:mm:ss");
    }
    
    void test_cell_formatted_as_date1()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        auto cell = ws.cell(xlnt::cell_reference(1, 1));
        
        cell.value(xlnt::datetime::today());
        cell.clear_value();
        TS_ASSERT(!cell.is_date()); // disagree with openpyxl
        TS_ASSERT(!cell.has_value());
    }
    
    void test_cell_formatted_as_date2()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        auto cell = ws.cell(xlnt::cell_reference(1, 1));
        
        cell.value(xlnt::datetime::today());
        cell.value("testme");
        TS_ASSERT(!cell.is_date());
        TS_ASSERT(cell.value<std::string>() == "testme");
    }
    
    void test_cell_formatted_as_date3()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        auto cell = ws.cell(xlnt::cell_reference(1, 1));
        
        cell.value(xlnt::datetime::today());
        cell.value(true);
        TS_ASSERT(!cell.is_date());
        TS_ASSERT(cell.value<bool>() == true);
    }
    
    void test_illegal_characters()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        auto cell = ws.cell(xlnt::cell_reference(1, 1));
    
        // The bytes 0x00 through 0x1F inclusive must be manually escaped in values.
        auto illegal_chrs = {0, 1, 2, 3, 4, 5, 6, 7, 8, 11, 12, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31};
        
        for(auto i : illegal_chrs)
        {
            std::string str(1, i);
            TS_ASSERT_THROWS(cell.value(str), xlnt::illegal_character);
        }
        
        cell.value(std::string(1, 33));
        cell.value(std::string(1, 9));  // Tab
        cell.value(std::string(1, 10));  // Newline
        cell.value(std::string(1, 13));  // Carriage return
        cell.value(" Leading and trailing spaces are legal ");
    }
    
    // void test_time_regex() {}
    
    void test_timedelta()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        auto cell = ws.cell(xlnt::cell_reference(1, 1));
        
        cell.value(xlnt::timedelta(1, 3, 0, 0, 0));

        TS_ASSERT(cell.value<long double>() == 1.125);
        TS_ASSERT(cell.data_type() == xlnt::cell::type::numeric);
        TS_ASSERT(!cell.is_date());
        TS_ASSERT(cell.number_format().format_string() == "[hh]:mm:ss");
    }

    void test_cell_offset()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        auto cell = ws.cell(xlnt::cell_reference(1, 1));
        TS_ASSERT(cell.offset(1, 2).reference() == "B3");
    }
    
    void test_font()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();    
        auto cell = ws.cell("A1");

		auto font = xlnt::font().bold(true);
        
        cell.font(font);
        
        TS_ASSERT(cell.has_format());
        TS_ASSERT(cell.format().font_applied());
        TS_ASSERT_EQUALS(cell.font(), font);
    }
    
    void test_fill()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        auto cell = ws.cell("A1");
        
		xlnt::fill fill(xlnt::pattern_fill()
			.type(xlnt::pattern_fill_type::solid)
			.foreground(xlnt::color::red()));
        cell.fill(fill);
        
        TS_ASSERT(cell.has_format());
        TS_ASSERT(cell.format().fill_applied());
        TS_ASSERT_EQUALS(cell.fill(), fill);
    }
    
    void test_border()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        auto cell = ws.cell("A1");
        
        xlnt::border border;
        cell.border(border);
        
        TS_ASSERT(cell.has_format());
        TS_ASSERT(cell.format().border_applied());
        TS_ASSERT_EQUALS(cell.border(), border);
    }
    
    void test_number_format()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        auto cell = ws.cell("A1");
        
        xlnt::number_format format("dd--hh--mm");
        cell.number_format(format);
        
        TS_ASSERT(cell.has_format());
        TS_ASSERT(cell.format().number_format_applied());
        TS_ASSERT_EQUALS(cell.number_format().format_string(), "dd--hh--mm");
    }
    
    void test_alignment()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        auto cell = ws.cell("A1");
        
        xlnt::alignment align;
        align.wrap(true);
    
        cell.alignment(align);
        
        TS_ASSERT(cell.has_format());
        TS_ASSERT(cell.format().alignment_applied());
        TS_ASSERT_EQUALS(cell.alignment(), align);
    }
    
    void test_protection()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        auto cell = ws.cell("A1");

        TS_ASSERT(!cell.has_format());

		auto protection = xlnt::protection().locked(false).hidden(true);
        cell.protection(protection);
        
        TS_ASSERT(cell.has_format());
        TS_ASSERT(cell.format().protection_applied());
        TS_ASSERT_EQUALS(cell.protection(), protection);
        
        TS_ASSERT(cell.has_format());
        cell.clear_format();
        TS_ASSERT(!cell.has_format());
    }
    
    void test_style()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        auto cell = ws.cell("A1");
        
        TS_ASSERT(!cell.has_style());
        
        auto test_style = wb.create_style("test_style");
        test_style.number_format(xlnt::number_format::date_ddmmyyyy(), true);
        
        cell.style(test_style);
        TS_ASSERT(cell.has_style());
        TS_ASSERT_EQUALS(cell.style().number_format(), xlnt::number_format::date_ddmmyyyy());
        TS_ASSERT_EQUALS(cell.style(), test_style);

        auto other_style = wb.create_style("other_style");
        other_style.number_format(xlnt::number_format::date_time2(), true);
        
        cell.style("other_style");
        TS_ASSERT_EQUALS(cell.style().number_format(), xlnt::number_format::date_time2());
        TS_ASSERT_EQUALS(cell.style(), other_style);
        
		auto last_style = wb.create_style("last_style");
        last_style.number_format(xlnt::number_format::percentage(), true);
        
        cell.style(last_style);
        TS_ASSERT_EQUALS(cell.style().number_format(), xlnt::number_format::percentage());
        TS_ASSERT_EQUALS(cell.style(), last_style);
        
        TS_ASSERT_THROWS(cell.style("doesn't exist"), xlnt::key_not_found);
        
        cell.clear_style();
        
        TS_ASSERT(!cell.has_style());
        TS_ASSERT_THROWS(cell.style(), xlnt::invalid_attribute);
    }
    
    void test_print()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();

        {
            auto cell = ws.cell("A1");

            std::stringstream ss;
            ss << cell;

            auto stream_string = ss.str();

            TS_ASSERT_EQUALS(cell.to_string(), stream_string);
            TS_ASSERT_EQUALS(stream_string, "");
        }

        {
            auto cell = ws.cell("A2");

            cell.value(false);

            std::stringstream ss;
            ss << cell;

            auto stream_string = ss.str();

            TS_ASSERT_EQUALS(cell.to_string(), stream_string);
            TS_ASSERT_EQUALS(stream_string, "FALSE");
        }
        
        {
            auto cell = ws.cell("A3");

            cell.value(true);

            std::stringstream ss;
            ss << cell;

            auto stream_string = ss.str();

            TS_ASSERT_EQUALS(cell.to_string(), stream_string);
            TS_ASSERT_EQUALS(stream_string, "TRUE");
        }
        
        {
            auto cell = ws.cell("A4");

            cell.value(1.234);

            std::stringstream ss;
            ss << cell;

            auto stream_string = ss.str();

            TS_ASSERT_EQUALS(cell.to_string(), stream_string);
            TS_ASSERT_EQUALS(stream_string, "1.234");
        }
        
        {
            auto cell = ws.cell("A5");

            cell.error("#REF");

            std::stringstream ss;
            ss << cell;

            auto stream_string = ss.str();

            TS_ASSERT_EQUALS(cell.to_string(), stream_string);
            TS_ASSERT_EQUALS(stream_string, "#REF");
        }
        
        {
            auto cell = ws.cell("A6");

            cell.value("test");

            std::stringstream ss;
            ss << cell;

            auto stream_string = ss.str();

            TS_ASSERT_EQUALS(cell.to_string(), stream_string);
            TS_ASSERT_EQUALS(stream_string, "test");
        }
    }

    void test_values()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        auto cell = ws.cell("A1");
        
        cell.value(static_cast<std::int8_t>(4));
        TS_ASSERT_EQUALS(cell.value<int8_t>(), 4);

        cell.value(static_cast<std::uint8_t>(3));
        TS_ASSERT_EQUALS(cell.value<std::uint8_t>(), 3);

        cell.value(static_cast<std::int16_t>(4));
        TS_ASSERT_EQUALS(cell.value<int16_t>(), 4);

        cell.value(static_cast<std::uint16_t>(3));
        TS_ASSERT_EQUALS(cell.value<std::uint16_t>(), 3);

        cell.value(static_cast<std::int32_t>(4));
        TS_ASSERT_EQUALS(cell.value<int32_t>(), 4);

        cell.value(static_cast<std::uint32_t>(3));
        TS_ASSERT_EQUALS(cell.value<std::uint32_t>(), 3);

        cell.value(static_cast<std::int64_t>(4));
        TS_ASSERT_EQUALS(cell.value<int64_t>(), 4);

        cell.value(static_cast<std::uint64_t>(3));
        TS_ASSERT_EQUALS(cell.value<std::uint64_t>(), 3);

        cell.value(static_cast<float>(3.14));
        TS_ASSERT_DELTA(cell.value<float>(), 3.14, 0.001);

        cell.value(static_cast<double>(4.1415));
        TS_ASSERT_EQUALS(cell.value<double>(), 4.1415);

        cell.value(static_cast<long double>(3.141592));
        TS_ASSERT_EQUALS(cell.value<long double>(), 3.141592);

        auto cell2 = ws.cell("A2");
        cell2.value(std::string(100'000, 'a'));
        cell.value(cell2);
        TS_ASSERT_EQUALS(cell.value<std::string>(), std::string(32'767, 'a'));
    }

    void test_reference()
    {
        xlnt::cell_reference_hash hash;
        TS_ASSERT_DIFFERS(hash(xlnt::cell_reference("A2")), hash(xlnt::cell_reference(1, 1)));
        TS_ASSERT_EQUALS(hash(xlnt::cell_reference("A2")), hash(xlnt::cell_reference(1, 2)));

        TS_ASSERT_EQUALS((xlnt::cell_reference("A1"), xlnt::cell_reference("B2")), xlnt::range_reference("A1:B2"));

        TS_ASSERT_THROWS(xlnt::cell_reference("A1&"), xlnt::invalid_cell_reference);
        TS_ASSERT_THROWS(xlnt::cell_reference("A"), xlnt::invalid_cell_reference);

        auto ref = xlnt::cell_reference("$B$7");
        TS_ASSERT(ref.row_absolute());
        TS_ASSERT(ref.column_absolute());

        TS_ASSERT(xlnt::cell_reference("A1") == "A1");
        TS_ASSERT(xlnt::cell_reference("A1") != "A2");
    }

    void test_anchor()
    {
        xlnt::workbook wb;
		auto cell = wb.active_sheet().cell("A1");
        auto anchor = cell.anchor();
        TS_ASSERT_EQUALS(anchor.first, 0);
        TS_ASSERT_EQUALS(anchor.second, 0);
    }

    void test_hyperlink()
    {
        xlnt::workbook wb;
		auto cell = wb.active_sheet().cell("A1");
        TS_ASSERT(!cell.has_hyperlink());
        TS_ASSERT_THROWS(cell.hyperlink(), xlnt::invalid_attribute);
        TS_ASSERT_THROWS(cell.hyperlink("notaurl"), xlnt::invalid_parameter);
        TS_ASSERT_THROWS(cell.hyperlink(""), xlnt::invalid_parameter);
        cell.hyperlink("http://example.com");
        TS_ASSERT(cell.has_hyperlink());
        TS_ASSERT_EQUALS(cell.hyperlink(), "http://example.com");
    }

    void test_comment()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        auto cell = ws.cell("A1");
        TS_ASSERT(!cell.has_comment());
        TS_ASSERT_THROWS(cell.comment(), xlnt::exception);
        cell.comment(xlnt::comment("comment", "author"));
        TS_ASSERT(cell.has_comment());
        TS_ASSERT_EQUALS(cell.comment(), xlnt::comment("comment", "author"));
        cell.clear_comment();
        TS_ASSERT(!cell.has_comment());
        TS_ASSERT_THROWS(cell.comment(), xlnt::exception);
    }
};
