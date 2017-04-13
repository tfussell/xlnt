// Copyright (c) 2014-2017 Thomas Fussell
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, WRISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE
//
// @license: http://www.opensource.org/licenses/mit-license.php
// @author: see AUTHORS file

#pragma once

#include <iostream>

#include <detail/vector_streambuf.hpp>
#include <detail/crypto/xlsx_crypto.hpp>
#include <helpers/temporary_file.hpp>
#include <helpers/test_suite.hpp>
#include <helpers/path_helper.hpp>
#include <helpers/xml_helper.hpp>
#include <xlnt/workbook/workbook.hpp>

class serialization_test_suite : public test_suite
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

    void test_non_xlsx()
    {
        xlnt::workbook wb;
        const auto path = path_helper::data_directory("1_powerpoint_presentation.xlsx");
        assert_throws(wb.load(path), xlnt::invalid_file);
    }

    void test_decrypt_agile()
    {
        xlnt::workbook wb;
        const auto path = path_helper::data_directory("5_encrypted_agile.xlsx");
        assert_throws_nothing(wb.load(path, "secret"));
    }

    void test_decrypt_libre_office()
    {
        xlnt::workbook wb;
        const auto path = path_helper::data_directory("6_encrypted_libre.xlsx");
        assert_throws_nothing(wb.load(path, u8"пароль"));
    }

    void test_decrypt_standard()
    {
        xlnt::workbook wb;
        const auto path = path_helper::data_directory("7_encrypted_standard.xlsx");
        assert_throws_nothing(wb.load(path, "password"));
    }

    void test_decrypt_numbers()
    {
        xlnt::workbook wb;
        const auto path = path_helper::data_directory("8_encrypted_numbers.xlsx");
        assert_throws_nothing(wb.load(path, "secret"));
    }

    void test_read_unicode_filename()
    {
#ifdef _MSC_VER
        xlnt::workbook wb;
        const auto path = LSTRING_LITERAL(XLNT_TEST_DATA_DIR) L"/9_unicode_Λ.xlsx";
        wb.load(path);
        assert_equals(wb.active_sheet().cell("A1").value<std::string>(), u8"unicodê!");
#endif

#ifndef __MINGW32__
        xlnt::workbook wb2;
        const auto path2 = U8STRING_LITERAL(XLNT_TEST_DATA_DIR) u8"/9_unicode_Λ.xlsx";
        wb2.load(path2);
        assert_equals(wb2.active_sheet().cell("A1").value<std::string>(), u8"unicodê!");
#endif
    }

    void test_comments()
    {
        xlnt::workbook wb;
        const auto path = path_helper::data_directory("10_comments_hyperlinks_formulae.xlsx");
        wb.load(path);

        auto sheet1 = wb[0];
        assert_equals(sheet1.cell("A1").value<std::string>(), "Sheet1!A1");
        assert_equals(sheet1.cell("A1").comment().plain_text(), "Sheet1 comment");
        assert_equals(sheet1.cell("A1").comment().author(), "Microsoft Office User");

        auto sheet2 = wb[1];
        assert_equals(sheet2.cell("A1").value<std::string>(), "Sheet2!A1");
        assert_equals(sheet2.cell("A1").comment().plain_text(), "Sheet2 comment");
        assert_equals(sheet2.cell("A1").comment().author(), "Microsoft Office User");
    }

    void test_read_hyperlink()
    {
        xlnt::workbook wb;
        const auto path = path_helper::data_directory("10_comments_hyperlinks_formulae.xlsx");
        wb.load(path);

        auto ws1 = wb.sheet_by_index(0);
        assert_equals(ws1.title(), "Sheet1");
        assert(ws1.cell("A4").has_hyperlink());
        assert_equals(ws1.cell("A4").value<std::string>(), "hyperlink1");
        assert_equals(ws1.cell("A4").hyperlink(), "https://microsoft.com/");
        assert(ws1.cell("A5").has_hyperlink());
        assert_equals(ws1.cell("A5").value<std::string>(), "https://google.com/");
        assert_equals(ws1.cell("A5").hyperlink(), "https://google.com/");
        //assert(ws1.cell("A6").has_hyperlink());
        assert_equals(ws1.cell("A6").value<std::string>(), "Sheet1!A1");
        //assert_equals(ws1.cell("A6").hyperlink(), "Sheet1!A1");
        assert(ws1.cell("A7").has_hyperlink());
        assert_equals(ws1.cell("A7").value<std::string>(), "mailto:invalid@example.com?subject=important");
        assert_equals(ws1.cell("A7").hyperlink(), "mailto:invalid@example.com?subject=important");

    }

    void test_read_formulae()
    {
        xlnt::workbook wb;
        const auto path = path_helper::data_directory("10_comments_hyperlinks_formulae.xlsx");
        wb.load(path);

        auto ws1 = wb.sheet_by_index(0);
        assert_equals(ws1.cell("C1").value<std::string>(), "ab");
        assert(ws1.cell("C1").has_formula());
        assert_equals(ws1.cell("C1").formula(), "CONCATENATE(C2,C3)");
        assert_equals(ws1.cell("C2").value<std::string>(), "a");
        assert_equals(ws1.cell("C3").value<std::string>(), "b");

        auto ws2 = wb.sheet_by_index(1);
        assert_equals(ws2.cell("C1").value<int>(), 6);
        assert(ws2.cell("C1").has_formula());
        assert_equals(ws2.cell("C1").formula(), "C2*C3");
        assert_equals(ws2.cell("C2").value<int>(), 2);
        assert_equals(ws2.cell("C3").value<int>(), 3);
    }

    void test_read_headers_and_footers()
    {
        xlnt::workbook wb;
        wb.load(path_helper::data_directory("11_print_settings.xlsx"));
        auto ws = wb.active_sheet();

        assert_equals(ws.cell("A1").value<std::string>(), "header");
        assert_equals(ws.cell("A2").value<std::string>(), "and");
        assert_equals(ws.cell("A3").value<std::string>(), "footer");
        assert_equals(ws.cell("A4").value<std::string>(), "page1");
        assert_equals(ws.cell("A43").value<std::string>(), "page2");

        assert(ws.has_header_footer());
        assert(ws.header_footer().align_with_margins());
        assert(ws.header_footer().scale_with_doc());
        assert(!ws.header_footer().different_first());
        assert(!ws.header_footer().different_odd_even());

        assert(ws.header_footer().has_header(xlnt::header_footer::location::left));
        assert_equals(ws.header_footer().header(xlnt::header_footer::location::left).plain_text(), "left header");
        assert(ws.header_footer().has_header(xlnt::header_footer::location::center));
        assert_equals(ws.header_footer().header(xlnt::header_footer::location::center).plain_text(), "center header");
        assert(ws.header_footer().has_header(xlnt::header_footer::location::right));
        assert_equals(ws.header_footer().header(xlnt::header_footer::location::right).plain_text(), "right header");
        assert(ws.header_footer().has_footer(xlnt::header_footer::location::left));
        assert_equals(ws.header_footer().footer(xlnt::header_footer::location::left).plain_text(), "left && footer");
        assert(ws.header_footer().has_footer(xlnt::header_footer::location::center));
        assert_equals(ws.header_footer().footer(xlnt::header_footer::location::center).plain_text(), "center footer");
        assert(ws.header_footer().has_footer(xlnt::header_footer::location::right));
        assert_equals(ws.header_footer().footer(xlnt::header_footer::location::right).plain_text(), "right footer");
    }

    void test_read_custom_properties()
    {
        xlnt::workbook wb;
        wb.load(path_helper::data_directory("12_advanced_properties.xlsx"));
        assert(wb.has_custom_property("Client"));
        assert_equals(wb.custom_property("Client").get<std::string>(), "me!");
    }

    /// <summary>
    /// Read file as an XLSX-formatted ZIP file in the filesystem to a workbook,
    /// write the workbook back to memory, then ensure that the contents of the two files are equivalent.
    /// </summary>
    bool round_trip_matches_rw(const xlnt::path &source)
    {
        xlnt::workbook source_workbook;
        source_workbook.load(source);

        std::vector<std::uint8_t> destination;
        source_workbook.save(destination);
        source_workbook.save("temp.xlsx");

#ifdef _MSC_VER
        std::ifstream source_stream(source.wstring(), std::ios::binary);
#else
        std::ifstream source_stream(source.string(), std::ios::binary);
#endif

        return xml_helper::xlsx_archives_match(xlnt::detail::to_vector(source_stream), destination);
    }

    bool round_trip_matches_rw(const xlnt::path &source, const std::string &password)
    {
        xlnt::workbook source_workbook;
        source_workbook.load(source, password);

        std::vector<std::uint8_t> destination;
        source_workbook.save(destination);

#ifdef _MSC_VER
        std::ifstream source_stream(source.wstring(), std::ios::binary);
#else
        std::ifstream source_stream(source.string(), std::ios::binary);
#endif

        const auto source_decrypted = xlnt::detail::decrypt_xlsx(
            xlnt::detail::to_vector(source_stream), password);

        return xml_helper::xlsx_archives_match(source_decrypted, destination);
    }

    void test_round_trip_rw()
    {
        const auto files = std::vector<std::string>
        {
            "2_minimal",
            "3_default",
            "4_every_style",
            u8"9_unicode_Λ",
            "10_comments_hyperlinks_formulae",
            "11_print_settings",
            "12_advanced_properties"
        };

        for (const auto file : files)
        {
            auto path = path_helper::data_directory(file + ".xlsx");
            assert(round_trip_matches_rw(path));
        }
    }

    void test_round_trip_rw_encrypted()
    {
        const auto files = std::vector<std::string>
        {
            "5_encrypted_agile",
            "6_encrypted_libre",
            "7_encrypted_standard",
            "8_encrypted_numbers"
        };

        for (const auto file : files)
        {
            auto path = path_helper::data_directory(file + ".xlsx");
            auto password = std::string(file == "7_encrypted_standard" ? "password"
                : file == "6_encrypted_libre" ? u8"пароль"
                : "secret");
            assert(round_trip_matches_rw(path, password));
        }
    }
};
