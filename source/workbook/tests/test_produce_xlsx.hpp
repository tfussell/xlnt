#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <detail/vector_streambuf.hpp>
#include <helpers/temporary_file.hpp>
#include <helpers/path_helper.hpp>
#include <helpers/xml_helper.hpp>
#include <xlnt/workbook/workbook.hpp>

class test_produce_xlsx : public CxxTest::TestSuite
{
public:
	bool workbook_matches_file(xlnt::workbook &wb, const xlnt::path &file)
	{
        std::vector<std::uint8_t> wb_data;
        wb.save(wb_data);

        std::ifstream file_stream(file.string(), std::ios::binary);
        std::vector<std::uint8_t> file_data;

        {
            xlnt::detail::vector_ostreambuf file_data_buffer(file_data);
            std::ostream file_data_stream(&file_data_buffer);
            file_data_stream << file_stream.rdbuf();
        }

		return xml_helper::xlsx_archives_match(wb_data, file_data);
	}

	void test_produce_minimal()
	{
		xlnt::workbook wb = xlnt::workbook::minimal();
		TS_ASSERT(workbook_matches_file(wb, path_helper::get_data_directory("8_minimal.xlsx")));
	}

	void test_produce_default_excel()
	{
		xlnt::workbook wb = xlnt::workbook::empty_excel();
		TS_ASSERT(workbook_matches_file(wb, path_helper::get_data_directory("9_default-excel.xlsx")));
	}

	void test_produce_default_libre_office()
	{
        TS_SKIP("");
		xlnt::workbook wb = xlnt::workbook::empty_libre_office();
		TS_ASSERT(workbook_matches_file(wb, path_helper::get_data_directory("10_default-libre-office.xlsx")));
	}

	void test_produce_simple_excel()
	{
		xlnt::workbook wb = xlnt::workbook::empty_excel();
		auto ws = wb.get_active_sheet();

		auto bold_font = xlnt::font().bold(true);

		ws.get_cell("A1").set_value("Type");
		ws.get_cell("A1").set_font(bold_font);

		ws.get_cell("B1").set_value("Value");
		ws.get_cell("B1").set_font(bold_font);

		ws.get_cell("A2").set_value("null");
		ws.get_cell("B2").set_value(nullptr);

		ws.get_cell("A3").set_value("bool (true)");
		ws.get_cell("B3").set_value(true);

		ws.get_cell("A4").set_value("bool (false)");
		ws.get_cell("B4").set_value(false);

		ws.get_cell("A5").set_value("number (std::int8_t)");
		ws.get_cell("B5").set_value(std::numeric_limits<std::int8_t>::max());

		ws.get_cell("A6").set_value("number (std::uint8_t)");
		ws.get_cell("B6").set_value(std::numeric_limits<std::uint8_t>::max());

		ws.get_cell("A7").set_value("number (std::uint16_t)");
		ws.get_cell("B7").set_value(std::numeric_limits<std::int16_t>::max());

		ws.get_cell("A8").set_value("number (std::uint16_t)");
		ws.get_cell("B8").set_value(std::numeric_limits<std::uint16_t>::max());

		ws.get_cell("A9").set_value("number (std::uint32_t)");
		ws.get_cell("B9").set_value(std::numeric_limits<std::int32_t>::max());

		ws.get_cell("A10").set_value("number (std::uint32_t)");
		ws.get_cell("B10").set_value(std::numeric_limits<std::uint32_t>::max());

		ws.get_cell("A11").set_value("number (std::uint64_t)");
		ws.get_cell("B11").set_value(std::numeric_limits<std::int64_t>::max());

		ws.get_cell("A12").set_value("number (std::uint64_t)");
		ws.get_cell("B12").set_value(std::numeric_limits<std::uint64_t>::max());

		ws.get_cell("A13").set_value("number (float)");
		ws.get_cell("B13").set_value(std::numeric_limits<float>::max());

		ws.get_cell("A14").set_value("number (double)");
		ws.get_cell("B14").set_value(std::numeric_limits<double>::max());

		ws.get_cell("A15").set_value("number (long double)");
		ws.get_cell("B15").set_value(std::numeric_limits<long double>::max());

		ws.get_cell("A16").set_value("text (char *)");
		ws.get_cell("B16").set_value("string");

		ws.get_cell("A17").set_value("text (std::string)");
		ws.get_cell("B17").set_value(std::string("string"));

		ws.get_cell("A18").set_value("date");
		ws.get_cell("B18").set_value(xlnt::date(2016, 2, 3));

		ws.get_cell("A19").set_value("time");
		ws.get_cell("B19").set_value(xlnt::time(1, 2, 3, 4));

		ws.get_cell("A20").set_value("datetime");
		ws.get_cell("B20").set_value(xlnt::datetime(2016, 2, 3, 1, 2, 3, 4));

		ws.get_cell("A21").set_value("timedelta");
		ws.get_cell("B21").set_value(xlnt::timedelta(1, 2, 3, 4, 5));

		ws.freeze_panes("B2");

		std::vector<std::uint8_t> temp_buffer;
		wb.save(temp_buffer);
		TS_ASSERT(!temp_buffer.empty());
	}

	void test_save_after_sheet_deletion()
	{
		xlnt::workbook workbook;

		TS_ASSERT_EQUALS(workbook.get_sheet_titles().size(), 1);

		auto sheet = workbook.create_sheet();
		sheet.set_title("XXX1");
		TS_ASSERT_EQUALS(workbook.get_sheet_titles().size(), 2);

		workbook.remove_sheet(workbook.get_sheet_by_title("XXX1"));
		TS_ASSERT_EQUALS(workbook.get_sheet_titles().size(), 1);

		std::vector<std::uint8_t> temp_buffer;
		TS_ASSERT_THROWS_NOTHING(workbook.save(temp_buffer));
		TS_ASSERT(!temp_buffer.empty());
	}

	void test_write_comments()
	{
		xlnt::workbook wb;

		xlnt::formatted_text comment_text;
		xlnt::text_run formatted_run;;
		formatted_run.set_bold(true);
		formatted_run.set_size(10);
		formatted_run.set_color(xlnt::indexed_color(81));
		formatted_run.set_font("Calibri");

		auto sheet1 = wb.get_active_sheet();

		sheet1.get_cell("A1").set_value("Sheet1!A1");
		formatted_run.set_string("Sheet1 comment");
		comment_text.add_run(formatted_run);
		sheet1.get_cell("A1").comment(xlnt::comment(comment_text, "Microsoft Office User"));

		sheet1.get_cell("A2").set_value("Sheet1!A2");
		formatted_run.set_string("Sheet1 comment2");
		comment_text.clear();
		comment_text.add_run(formatted_run);
		sheet1.get_cell("A2").comment(xlnt::comment(comment_text, "Microsoft Office User"));

		auto sheet2 = wb.create_sheet();
		sheet2.get_cell("A1").set_value("Sheet2!A1");
		formatted_run.set_string("Sheet2 comment");
		comment_text.clear();
		comment_text.add_run(formatted_run);
		sheet2.get_cell("A1").comment(xlnt::comment(comment_text, "Microsoft Office User"));

		sheet2.get_cell("A2").set_value("Sheet2!A2");
		formatted_run.set_string("Sheet2 comment2");
		comment_text.clear();
		comment_text.add_run(formatted_run);
		sheet2.get_cell("A2").comment(xlnt::comment(comment_text, "Microsoft Office User"));

		wb.save("debug.xlsx");

		//TS_ASSERT(workbook_matches_file(wb, xlnt::path("data/18_basic_comments.xlsx")));
	}
};
