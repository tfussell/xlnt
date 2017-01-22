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

	void test_produce_empty()
	{
		xlnt::workbook wb;
        const auto path = path_helper::data_directory("3_default.xlsx");
		TS_ASSERT(workbook_matches_file(wb, path));
	}

	void test_produce_simple_excel()
	{
		xlnt::workbook wb;
		auto ws = wb.active_sheet();

		auto bold_font = xlnt::font().bold(true);

		ws.cell("A1").value("Type");
		ws.cell("A1").font(bold_font);

		ws.cell("B1").value("Value");
		ws.cell("B1").font(bold_font);

		ws.cell("A2").value("null");
		ws.cell("B2").value(nullptr);

		ws.cell("A3").value("bool (true)");
		ws.cell("B3").value(true);

		ws.cell("A4").value("bool (false)");
		ws.cell("B4").value(false);

		ws.cell("A5").value("number (std::int8_t)");
		ws.cell("B5").value(std::numeric_limits<std::int8_t>::max());

		ws.cell("A6").value("number (std::uint8_t)");
		ws.cell("B6").value(std::numeric_limits<std::uint8_t>::max());

		ws.cell("A7").value("number (std::uint16_t)");
		ws.cell("B7").value(std::numeric_limits<std::int16_t>::max());

		ws.cell("A8").value("number (std::uint16_t)");
		ws.cell("B8").value(std::numeric_limits<std::uint16_t>::max());

		ws.cell("A9").value("number (std::uint32_t)");
		ws.cell("B9").value(std::numeric_limits<std::int32_t>::max());

		ws.cell("A10").value("number (std::uint32_t)");
		ws.cell("B10").value(std::numeric_limits<std::uint32_t>::max());

		ws.cell("A11").value("number (std::uint64_t)");
		ws.cell("B11").value(std::numeric_limits<std::int64_t>::max());

		ws.cell("A12").value("number (std::uint64_t)");
		ws.cell("B12").value(std::numeric_limits<std::uint64_t>::max());

		ws.cell("A13").value("number (float)");
		ws.cell("B13").value(std::numeric_limits<float>::max());

		ws.cell("A14").value("number (double)");
		ws.cell("B14").value(std::numeric_limits<double>::max());

		ws.cell("A15").value("number (long double)");
		ws.cell("B15").value(std::numeric_limits<long double>::max());

		ws.cell("A16").value("text (char *)");
		ws.cell("B16").value("string");

		ws.cell("A17").value("text (std::string)");
		ws.cell("B17").value(std::string("string"));

		ws.cell("A18").value("date");
		ws.cell("B18").value(xlnt::date(2016, 2, 3));

		ws.cell("A19").value("time");
		ws.cell("B19").value(xlnt::time(1, 2, 3, 4));

		ws.cell("A20").value("datetime");
		ws.cell("B20").value(xlnt::datetime(2016, 2, 3, 1, 2, 3, 4));

		ws.cell("A21").value("timedelta");
		ws.cell("B21").value(xlnt::timedelta(1, 2, 3, 4, 5));

		ws.freeze_panes("B2");

		std::vector<std::uint8_t> temp_buffer;
		wb.save(temp_buffer);
		TS_ASSERT(!temp_buffer.empty());
	}

	void test_save_after_sheet_deletion()
	{
		xlnt::workbook workbook;

		TS_ASSERT_EQUALS(workbook.sheet_titles().size(), 1);

		auto sheet = workbook.create_sheet();
		sheet.title("XXX1");
		TS_ASSERT_EQUALS(workbook.sheet_titles().size(), 2);

		workbook.remove_sheet(workbook.sheet_by_title("XXX1"));
		TS_ASSERT_EQUALS(workbook.sheet_titles().size(), 1);

		std::vector<std::uint8_t> temp_buffer;
		TS_ASSERT_THROWS_NOTHING(workbook.save(temp_buffer));
		TS_ASSERT(!temp_buffer.empty());
	}

	void test_write_comments()
	{
		xlnt::workbook wb;
		auto sheet1 = wb.active_sheet();
        auto comment_font = xlnt::font().bold(true).size(10).color(xlnt::indexed_color(81)).name("Calibri");

		sheet1.cell("A1").value("Sheet1!A1");
		sheet1.cell("A1").comment("Sheet1 comment", comment_font, "Microsoft Office User");

		sheet1.cell("A2").value("Sheet1!A2");
		sheet1.cell("A2").comment("Sheet1 comment2", comment_font, "Microsoft Office User");

		auto sheet2 = wb.create_sheet();
		sheet2.cell("A1").value("Sheet2!A1");
		sheet2.cell("A2").comment("Sheet2 comment", comment_font, "Microsoft Office User");

		sheet2.cell("A2").value("Sheet2!A2");
		sheet2.cell("A2").comment("Sheet2 comment2", comment_font, "Microsoft Office User");

        const auto path = path_helper::data_directory("10_comments_hyperlinks_formulae.xlsx");
		TS_ASSERT(workbook_matches_file(wb, path));
	}

    void test_save_after_clear_all_formulae()
    {
        xlnt::workbook wb;
        const auto path = path_helper::data_directory("10_comments_hyperlinks_formulae.xlsx");
        wb.load(path);

        auto ws1 = wb.sheet_by_index(0);
        TS_ASSERT(ws1.cell("C1").has_formula());
        TS_ASSERT_EQUALS(ws1.cell("C1").formula(), "CONCATENATE(C2,C3)");
        ws1.cell("C1").clear_formula();

        auto ws2 = wb.sheet_by_index(1);
        TS_ASSERT(ws2.cell("C1").has_formula());
        TS_ASSERT_EQUALS(ws2.cell("C1").formula(), "C2*C3");
        ws2.cell("C1").clear_formula();

        wb.save("clear_formulae.xlsx");
    }
};
