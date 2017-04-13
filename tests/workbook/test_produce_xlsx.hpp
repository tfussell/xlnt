#pragma once

#include <iostream>
#include <helpers/test_suite.hpp>

#include <detail/vector_streambuf.hpp>
#include <helpers/temporary_file.hpp>
#include <helpers/path_helper.hpp>
#include <helpers/xml_helper.hpp>
#include <xlnt/workbook/workbook.hpp>

class test_produce_xlsx : public test_suite
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
		assert(workbook_matches_file(wb, path));
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

		ws.cell("A5").value("number (int)");
		ws.cell("B5").value(std::numeric_limits<int>::max());

		ws.cell("A5").value("number (unsigned int)");
		ws.cell("B5").value(std::numeric_limits<unsigned int>::max());

		ws.cell("A6").value("number (long long int)");
		ws.cell("B6").value(std::numeric_limits<long long int>::max());

		ws.cell("A6").value("number (unsigned long long int)");
		ws.cell("B6").value(std::numeric_limits<unsigned long long int>::max());

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
		assert(!temp_buffer.empty());
	}

	void test_save_after_sheet_deletion()
	{
		xlnt::workbook workbook;

		assert_equals(workbook.sheet_titles().size(), 1);

		auto sheet = workbook.create_sheet();
		sheet.title("XXX1");
		assert_equals(workbook.sheet_titles().size(), 2);

		workbook.remove_sheet(workbook.sheet_by_title("XXX1"));
		assert_equals(workbook.sheet_titles().size(), 1);

		std::vector<std::uint8_t> temp_buffer;
		assert_throws_nothing(workbook.save(temp_buffer));
		assert(!temp_buffer.empty());
	}

	void test_write_comments_hyperlinks_formulae()
	{
		xlnt::workbook wb;
		auto sheet1 = wb.active_sheet();
        auto comment_font = xlnt::font().bold(true).size(10).color(xlnt::indexed_color(81)).name("Calibri");

		sheet1.cell("A1").value("Sheet1!A1");
		sheet1.cell("A1").comment("Sheet1 comment", comment_font, "Microsoft Office User");

		sheet1.cell("A2").value("Sheet1!A2");
		sheet1.cell("A2").comment("Sheet1 comment2", comment_font, "Microsoft Office User");

        sheet1.cell("A4").hyperlink("https://microsoft.com", "hyperlink1");
        sheet1.cell("A5").hyperlink("https://google.com");
        sheet1.cell("A6").hyperlink(sheet1.cell("A1"));
        sheet1.cell("A7").hyperlink("mailto:invalid@example.com?subject=important");

        sheet1.cell("C1").formula("=CONCATENATE(C2,C3)");
        sheet1.cell("C2").value("a");
        sheet1.cell("C3").value("b");

		auto sheet2 = wb.create_sheet();
		sheet2.cell("A1").value("Sheet2!A1");
		sheet2.cell("A2").comment("Sheet2 comment", comment_font, "Microsoft Office User");

		sheet2.cell("A2").value("Sheet2!A2");
		sheet2.cell("A2").comment("Sheet2 comment2", comment_font, "Microsoft Office User");

        sheet2.cell("A4").hyperlink("https://apple.com", "hyperlink2");

        sheet2.cell("C1").formula("=C2*C3");
        sheet2.cell("C2").value(2);
        sheet2.cell("C3").value(3);

        const auto path = path_helper::data_directory("10_comments_hyperlinks_formulae.xlsx");
		assert(workbook_matches_file(wb, path));
	}

    void test_save_after_clear_all_formulae()
    {
        xlnt::workbook wb;
        const auto path = path_helper::data_directory("10_comments_hyperlinks_formulae.xlsx");
        wb.load(path);

        auto ws1 = wb.sheet_by_index(0);
        assert(ws1.cell("C1").has_formula());
        assert_equals(ws1.cell("C1").formula(), "CONCATENATE(C2,C3)");
        ws1.cell("C1").clear_formula();

        auto ws2 = wb.sheet_by_index(1);
        assert(ws2.cell("C1").has_formula());
        assert_equals(ws2.cell("C1").formula(), "C2*C3");
        ws2.cell("C1").clear_formula();

        wb.save("clear_formulae.xlsx");
    }
};
