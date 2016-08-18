#pragma once

#include <ctime>
#include <iostream>
#include <sstream>
#include <cxxtest/TestSuite.h>

#include <xlnt/xlnt.hpp>

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

		cell.set_value("03:");
        TS_ASSERT_EQUALS(cell.get_value<std::string>(), "03:");

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
    }

    void test_null()
    {
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
        auto ws = wb_guess_types.create_sheet();
        auto cell = ws.get_cell(xlnt::cell_reference(1, 1));
        
        cell.set_value("=if(A1<4;-1;1)");
        TS_ASSERT(cell.get_data_type() == xlnt::cell::type::formula);
    }

    void test_formula3()
    {
        auto ws = wb.create_sheet();
        auto cell = ws.get_cell(xlnt::cell_reference(1, 1));

        TS_ASSERT(!cell.has_formula());
        TS_ASSERT_THROWS(cell.set_formula(""), xlnt::invalid_parameter);
        TS_ASSERT_THROWS(cell.get_formula(), xlnt::invalid_attribute);
        cell.set_formula("=42");
        TS_ASSERT(cell.has_formula());
        TS_ASSERT_EQUALS(cell.get_formula(), "42");
        cell.clear_formula();
        TS_ASSERT(!cell.has_formula());
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
    
    void test_illegal_characters()
    {
        auto ws = wb.create_sheet();
        auto cell = ws.get_cell(xlnt::cell_reference(1, 1));
    
        // The bytes 0x00 through 0x1F inclusive must be manually escaped in values.
        auto illegal_chrs = {0, 1, 2, 3, 4, 5, 6, 7, 8, 11, 12, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31};
        
        for(auto i : illegal_chrs)
        {
            std::string str(1, i);
            TS_ASSERT_THROWS(cell.set_value(str), xlnt::illegal_character);
        }
        
        cell.set_value(std::string(1, 33));
        cell.set_value(std::string(1, 9));  // Tab
        cell.set_value(std::string(1, 10));  // Newline
        cell.set_value(std::string(1, 13));  // Carriage return
        cell.set_value(" Leading and trailing spaces are legal ");
    }
    
    // void test_time_regex() {}
    
    void test_timedelta()
    {
        auto ws = wb.create_sheet();
        auto cell = ws.get_cell(xlnt::cell_reference(1, 1));
        
        cell.set_value(xlnt::timedelta(1, 3, 0, 0, 0));

        TS_ASSERT(cell.get_value<long double>() == 1.125);
        TS_ASSERT(cell.get_data_type() == xlnt::cell::type::numeric);
        TS_ASSERT(!cell.is_date());
        TS_ASSERT(cell.get_number_format().get_format_string() == "[hh]:mm:ss");
    }
    
    void test_repr()
    {
        auto ws = wb[1];
        auto cell = ws.get_cell(xlnt::cell_reference(1, 1));
        
        TS_ASSERT(cell.to_repr() == "<Cell Sheet2.A1>");
    }
    
    void test_cell_offset()
    {
        auto ws = wb.create_sheet();
        auto cell = ws.get_cell(xlnt::cell_reference(1, 1));
        TS_ASSERT(cell.offset(1, 2).get_reference() == "B3");
    }
    
    void test_font()
    {
        auto ws = wb.create_sheet();    
        auto cell = ws.get_cell("A1");

		auto font = xlnt::font().bold(true);
        
        cell.set_font(font);
        
        TS_ASSERT(cell.has_format());
        TS_ASSERT(cell.get_format().font_applied());
        TS_ASSERT_EQUALS(cell.get_font(), font);
    }
    
    void test_fill()
    {
        auto ws = wb.create_sheet();
        auto cell = ws.get_cell("A1");
        
		xlnt::fill fill(xlnt::pattern_fill()
			.type(xlnt::pattern_fill_type::solid)
			.foreground(xlnt::color::red()));
        cell.set_fill(fill);
        
        TS_ASSERT(cell.has_format());
        TS_ASSERT(cell.get_format().fill_applied());
        TS_ASSERT_EQUALS(cell.get_fill(), fill);
    }
    
    void test_border()
    {
        auto ws = wb.create_sheet();
        auto cell = ws.get_cell("A1");
        
        xlnt::border border;

        cell.set_border(border);
        
        TS_ASSERT(cell.has_format());
        TS_ASSERT(cell.get_format().border_applied());
        TS_ASSERT_EQUALS(cell.get_border(), border);
    }
    
    void test_number_format()
    {
        auto ws = wb.create_sheet();
        auto cell = ws.get_cell("A1");
        
        xlnt::number_format format("dd--hh--mm");
        cell.set_number_format(format);
        
        TS_ASSERT(cell.has_format());
        TS_ASSERT(cell.get_format().number_format_applied());
        TS_ASSERT_EQUALS(cell.get_number_format().get_format_string(), "dd--hh--mm");
    }
    
    void test_alignment()
    {
        auto ws = wb.create_sheet();
        auto cell = ws.get_cell("A1");
        
        xlnt::alignment align;
        align.wrap(true);
    
        cell.set_alignment(align);
        
        TS_ASSERT(cell.has_format());
        TS_ASSERT(cell.get_format().alignment_applied());
        TS_ASSERT_EQUALS(cell.get_alignment(), align);
    }
    
    void test_protection()
    {
        auto ws = wb.create_sheet();
        auto cell = ws.get_cell("A1");

        TS_ASSERT(!cell.has_format());

		auto protection = xlnt::protection().locked(false).hidden(true);
        cell.set_protection(protection);
        
        TS_ASSERT(cell.has_format());
        TS_ASSERT(cell.get_format().protection_applied());
        TS_ASSERT_EQUALS(cell.get_protection(), protection);
        
        TS_ASSERT(cell.has_format());
        cell.clear_format();
        TS_ASSERT(!cell.has_format());
    }
    
    void test_style()
    {
        auto ws = wb.create_sheet();
        auto cell = ws.get_cell("A1");
        
        TS_ASSERT(!cell.has_style());
        
        auto &test_style = wb.create_style("test_style");
        test_style.number_format(xlnt::number_format::date_ddmmyyyy(), true);
        
        cell.set_style(test_style);
        TS_ASSERT(cell.has_style());
        TS_ASSERT_EQUALS(cell.get_style().number_format(), xlnt::number_format::date_ddmmyyyy());
        TS_ASSERT_EQUALS(cell.get_style(), test_style);

        auto &other_style = wb.create_style("other_style");
        other_style.number_format(xlnt::number_format::date_time2(), true);
        
        cell.set_style("other_style");
        TS_ASSERT_EQUALS(cell.get_style().number_format(), xlnt::number_format::date_time2());
        TS_ASSERT_EQUALS(cell.get_style(), other_style);
        
		auto &last_style = wb.create_style("last_style");
        last_style.number_format(xlnt::number_format::percentage(), true);
        
        cell.set_style(last_style);
        TS_ASSERT_EQUALS(cell.get_style().number_format(), xlnt::number_format::percentage());
        TS_ASSERT_EQUALS(cell.get_style(), last_style);
        
        TS_ASSERT_THROWS(cell.set_style("doesn't exist"), xlnt::key_not_found);
        
        cell.clear_style();
        
        TS_ASSERT(!cell.has_style());
        TS_ASSERT_THROWS(cell.get_style(), xlnt::invalid_attribute);
    }
    
    void test_print()
    {
        auto ws = wb.create_sheet();

        {
            auto cell = ws.get_cell("A1");

            std::stringstream ss;
            ss << cell;

            auto stream_string = ss.str();

            TS_ASSERT_EQUALS(cell.to_string(), stream_string);
            TS_ASSERT_EQUALS(stream_string, "");
        }

        {
            auto cell = ws.get_cell("A2");

            cell.set_value(false);

            std::stringstream ss;
            ss << cell;

            auto stream_string = ss.str();

            TS_ASSERT_EQUALS(cell.to_string(), stream_string);
            TS_ASSERT_EQUALS(stream_string, "FALSE");
        }
        
        {
            auto cell = ws.get_cell("A3");

            cell.set_value(true);

            std::stringstream ss;
            ss << cell;

            auto stream_string = ss.str();

            TS_ASSERT_EQUALS(cell.to_string(), stream_string);
            TS_ASSERT_EQUALS(stream_string, "TRUE");
        }
        
        {
            auto cell = ws.get_cell("A4");

            cell.set_value(1.234);

            std::stringstream ss;
            ss << cell;

            auto stream_string = ss.str();

            TS_ASSERT_EQUALS(cell.to_string(), stream_string);
            TS_ASSERT_EQUALS(stream_string, "1.234");
        }
        
        {
            auto cell = ws.get_cell("A5");

            cell.set_error("#REF");

            std::stringstream ss;
            ss << cell;

            auto stream_string = ss.str();

            TS_ASSERT_EQUALS(cell.to_string(), stream_string);
            TS_ASSERT_EQUALS(stream_string, "#REF");
        }
        
        {
            auto cell = ws.get_cell("A6");

            cell.set_value("test");

            std::stringstream ss;
            ss << cell;

            auto stream_string = ss.str();

            TS_ASSERT_EQUALS(cell.to_string(), stream_string);
            TS_ASSERT_EQUALS(stream_string, "test");
        }
    }

    void test_values()
    {
        auto ws = wb.create_sheet();
        auto cell = ws.get_cell("A1");
        
        cell.set_value(static_cast<std::int8_t>(4));
        TS_ASSERT_EQUALS(cell.get_value<int8_t>(), 4);

        cell.set_value(static_cast<std::uint8_t>(3));
        TS_ASSERT_EQUALS(cell.get_value<std::uint8_t>(), 3);

        cell.set_value(static_cast<std::int16_t>(4));
        TS_ASSERT_EQUALS(cell.get_value<int16_t>(), 4);

        cell.set_value(static_cast<std::uint16_t>(3));
        TS_ASSERT_EQUALS(cell.get_value<std::uint16_t>(), 3);

        cell.set_value(static_cast<std::int32_t>(4));
        TS_ASSERT_EQUALS(cell.get_value<int32_t>(), 4);

        cell.set_value(static_cast<std::uint32_t>(3));
        TS_ASSERT_EQUALS(cell.get_value<std::uint32_t>(), 3);

        cell.set_value(static_cast<std::int64_t>(4));
        TS_ASSERT_EQUALS(cell.get_value<int64_t>(), 4);

        cell.set_value(static_cast<std::uint64_t>(3));
        TS_ASSERT_EQUALS(cell.get_value<std::uint64_t>(), 3);

#ifdef __linux
        cell.set_value(static_cast<long long>(3));
        TS_ASSERT_EQUALS(cell.get_value<long long>(), 3);

        cell.set_value(static_cast<unsigned long long>(3));
        TS_ASSERT_EQUALS(cell.get_value<unsigned long long>(), 3);
#endif

        cell.set_value(static_cast<std::uint64_t>(3));
        TS_ASSERT_EQUALS(cell.get_value<std::uint64_t>(), 3);

        cell.set_value(static_cast<float>(3.14));
        TS_ASSERT_DELTA(cell.get_value<float>(), 3.14, 0.001);

        cell.set_value(static_cast<double>(4.1415));
        TS_ASSERT_EQUALS(cell.get_value<double>(), 4.1415);

        cell.set_value(static_cast<long double>(3.141592));
        TS_ASSERT_EQUALS(cell.get_value<long double>(), 3.141592);

        auto cell2 = ws.get_cell("A2");
        cell2.set_value(std::string(100'000, 'a'));
        cell.set_value(cell2);
        TS_ASSERT_EQUALS(cell.get_value<std::string>(), std::string(32'767, 'a'));
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
		auto cell = wb.get_active_sheet().get_cell("A1");
        auto anchor = cell.get_anchor();
        TS_ASSERT_EQUALS(anchor.first, 0);
        TS_ASSERT_EQUALS(anchor.second, 0);
    }

    void test_hyperlink()
    {
        xlnt::workbook wb;
		auto cell = wb.get_active_sheet().get_cell("A1");
        TS_ASSERT(!cell.has_hyperlink());
        TS_ASSERT_THROWS(cell.get_hyperlink(), xlnt::invalid_attribute);
        TS_ASSERT_THROWS(cell.set_hyperlink("notaurl"), xlnt::invalid_parameter);
        TS_ASSERT_THROWS(cell.set_hyperlink(""), xlnt::invalid_parameter);
        cell.set_hyperlink("http://example.com");
        TS_ASSERT(cell.has_hyperlink());
        TS_ASSERT_EQUALS(cell.get_hyperlink(), "http://example.com");
    }
};
