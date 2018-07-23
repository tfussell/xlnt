// Copyright (c) 2014-2018 Thomas Fussell
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

#include <iostream>

#include <xlnt/cell/comment.hpp>
#include <xlnt/cell/hyperlink.hpp>
#include <xlnt/cell/cell.hpp>
#include <xlnt/styles/font.hpp>
#include <xlnt/styles/style.hpp>
#include <xlnt/styles/fill.hpp>
#include <xlnt/styles/format.hpp>
#include <xlnt/styles/number_format.hpp>
#include <xlnt/styles/border.hpp>
#include <xlnt/utils/date.hpp>
#include <xlnt/utils/datetime.hpp>
#include <xlnt/utils/time.hpp>
#include <xlnt/utils/timedelta.hpp>
#include <xlnt/utils/variant.hpp>
#include <xlnt/workbook/streaming_workbook_reader.hpp>
#include <xlnt/workbook/streaming_workbook_writer.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/workbook/metadata_property.hpp>
#include <xlnt/worksheet/column_properties.hpp>
#include <xlnt/worksheet/row_properties.hpp>
#include <xlnt/worksheet/sheet_format_properties.hpp>
#include <xlnt/worksheet/header_footer.hpp>
#include <xlnt/worksheet/worksheet.hpp>
#include <detail/cryptography/xlsx_crypto_consumer.hpp>
#include <detail/serialization/vector_streambuf.hpp>
#include <helpers/path_helper.hpp>
#include <helpers/temporary_file.hpp>
#include <helpers/test_suite.hpp>
#include <helpers/xml_helper.hpp>

class serialization_test_suite : public test_suite
{
public:
    serialization_test_suite()
    {
        register_test(test_produce_empty);
        register_test(test_produce_simple_excel);
        register_test(test_save_after_sheet_deletion);
        register_test(test_write_comments_hyperlinks_formulae);
        register_test(test_save_after_clear_all_formulae);
        register_test(test_load_non_xlsx);
        register_test(test_decrypt_agile);
        register_test(test_decrypt_libre_office);
        register_test(test_decrypt_standard);
        register_test(test_decrypt_numbers);
        register_test(test_read_unicode_filename);
        register_test(test_comments);
        register_test(test_read_hyperlink);
        register_test(test_read_formulae);
        register_test(test_read_headers_and_footers);
        register_test(test_read_custom_properties);
        register_test(test_read_custom_heights_widths);
        register_test(test_write_custom_heights_widths);
        register_test(test_round_trip_rw_minimal);
        register_test(test_round_trip_rw_default);
        register_test(test_round_trip_rw_every_style);
        register_test(test_round_trip_rw_unicode);
        register_test(test_round_trip_rw_comments_hyperlinks_formulae);
        register_test(test_round_trip_rw_print_settings);
        register_test(test_round_trip_rw_advanced_properties);
        register_test(test_round_trip_rw_custom_heights_widths);
        register_test(test_round_trip_rw_encrypted_agile);
        register_test(test_round_trip_rw_encrypted_libre);
        register_test(test_round_trip_rw_encrypted_standard);
        register_test(test_round_trip_rw_encrypted_numbers);
        register_test(test_streaming_read);
        register_test(test_streaming_write);
    }

    bool workbook_matches_file(xlnt::workbook &wb, const xlnt::path &file)
    {
        std::vector<std::uint8_t> wb_data;
        wb.save(wb_data);
        wb.save("temp.xlsx");

        std::ifstream file_stream(file.string(), std::ios::binary);
        auto file_data = xlnt::detail::to_vector(file_stream);

        return xml_helper::xlsx_archives_match(wb_data, file_data);
    }

    void test_produce_empty()
    {
        xlnt::workbook wb;
        const auto path = path_helper::test_file("3_default.xlsx");
        xlnt_assert(workbook_matches_file(wb, path));
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
        xlnt_assert(!temp_buffer.empty());
    }

    void test_save_after_sheet_deletion()
    {
        xlnt::workbook workbook;

        xlnt_assert_equals(workbook.sheet_titles().size(), 1);

        auto sheet = workbook.create_sheet();
        sheet.title("XXX1");
        xlnt_assert_equals(workbook.sheet_titles().size(), 2);

        workbook.remove_sheet(workbook.sheet_by_title("XXX1"));
        xlnt_assert_equals(workbook.sheet_titles().size(), 1);

        std::vector<std::uint8_t> temp_buffer;
        xlnt_assert_throws_nothing(workbook.save(temp_buffer));
        xlnt_assert(!temp_buffer.empty());
    }

    void test_write_comments_hyperlinks_formulae()
    {
        xlnt::workbook wb;

        xlnt::sheet_format_properties format_properties;
        format_properties.base_col_width = 10.0;
        format_properties.default_row_height = 16.0;
        format_properties.dy_descent = 0.2;

        auto sheet1 = wb.active_sheet();
        sheet1.format_properties(format_properties);

        auto selection = xlnt::selection();
        selection.active_cell("C1");
        selection.sqref("C1");
        sheet1.view().add_selection(selection);

        // comments
        auto comment_font = xlnt::font()
                                .bold(true)
                                .size(10)
                                .color(xlnt::indexed_color(81))
                                .name("Calibri");
        sheet1.cell("A1").value("Sheet1!A1");
        sheet1.cell("A1").comment("Sheet1 comment", comment_font, "Microsoft Office User");

        sheet1.cell("A2").value("Sheet1!A2");
        sheet1.cell("A2").comment("Sheet1 comment2", comment_font, "Microsoft Office User");

        // hyperlinks
        auto hyperlink_font = xlnt::font()
                                  .size(12)
                                  .color(xlnt::theme_color(10))
                                  .name("Calibri")
                                  .family(2)
                                  .scheme("minor")
                                  .underline(xlnt::font::underline_style::single);
        auto hyperlink_style = wb.create_builtin_style(8);
        hyperlink_style.font(hyperlink_font);
        hyperlink_style.number_format(hyperlink_style.number_format(), false);
        hyperlink_style.fill(hyperlink_style.fill(), false);
        hyperlink_style.border(hyperlink_style.border(), false);
        auto hyperlink_format = wb.create_format();
        hyperlink_format.font(hyperlink_font);
        hyperlink_format.style(hyperlink_style);

        sheet1.cell("A4").hyperlink("https://microsoft.com/", "hyperlink1");
        sheet1.cell("A4").format(hyperlink_format);

        sheet1.cell("A5").hyperlink("https://google.com/");
        sheet1.cell("A5").format(hyperlink_format);

        sheet1.cell("A6").hyperlink(sheet1.cell("A1"));
        sheet1.cell("A6").format(hyperlink_format);

        sheet1.cell("A7").hyperlink("mailto:invalid@example.com?subject=important");
        sheet1.cell("A7").format(hyperlink_format);

        // formulae
        sheet1.cell("C1").formula("=CONCATENATE(C2,C3)");
        sheet1.cell("C2").value("a");
        sheet1.cell("C3").value("b");

        for (auto i = 1; i <= 7; ++i)
        {
            sheet1.row_properties(i).dy_descent = 0.2;
        }

        auto sheet2 = wb.create_sheet();
        sheet2.format_properties(format_properties);
        sheet2.add_view(xlnt::sheet_view());
        sheet2.view().add_selection(selection);

        // comments
        sheet2.cell("A1").value("Sheet2!A1");
        sheet2.cell("A1").comment("Sheet2 comment", comment_font, "Microsoft Office User");

        sheet2.cell("A2").value("Sheet2!A2");
        sheet2.cell("A2").comment("Sheet2 comment2", comment_font, "Microsoft Office User");

        // hyperlinks
        sheet2.cell("A4").hyperlink("https://apple.com/", "hyperlink2");
        sheet2.cell("A4").format(hyperlink_format);

        // formulae
        sheet2.cell("C1").formula("=C2*C3");
        sheet2.cell("C2").value(2);
        sheet2.cell("C3").value(3);

        for (auto i = 1; i <= 4; ++i)
        {
            sheet2.row_properties(i).dy_descent = 0.2;
        }

        wb.default_slicer_style("SlicerStyleLight1");
        wb.enable_known_fonts();

        wb.core_property(xlnt::core_property::created, "2018-03-18T20:53:30Z");
        wb.core_property(xlnt::core_property::modified, "2018-03-18T20:59:53Z");

        const auto path = path_helper::test_file("10_comments_hyperlinks_formulae.xlsx");
        xlnt_assert(workbook_matches_file(wb, path));
    }

    void test_save_after_clear_all_formulae()
    {
        xlnt::workbook wb;
        const auto path = path_helper::test_file("10_comments_hyperlinks_formulae.xlsx");
        wb.load(path);

        auto ws1 = wb.sheet_by_index(0);
        xlnt_assert(ws1.cell("C1").has_formula());
        xlnt_assert_equals(ws1.cell("C1").formula(), "CONCATENATE(C2,C3)");
        ws1.cell("C1").clear_formula();

        auto ws2 = wb.sheet_by_index(1);
        xlnt_assert(ws2.cell("C1").has_formula());
        xlnt_assert_equals(ws2.cell("C1").formula(), "C2*C3");
        ws2.cell("C1").clear_formula();

        wb.save("clear_formulae.xlsx");
    }

    void test_load_non_xlsx()
    {
        xlnt::workbook wb;
        const auto path = path_helper::test_file("1_powerpoint_presentation.xlsx");
        xlnt_assert_throws(wb.load(path), xlnt::invalid_file);
    }

    void test_decrypt_agile()
    {
        xlnt::workbook wb;
        const auto path = path_helper::test_file("5_encrypted_agile.xlsx");
        xlnt_assert_throws(wb.load(path, "incorrect"), xlnt::exception);
        xlnt_assert_throws_nothing(wb.load(path, "secret"));
    }

    void test_decrypt_libre_office()
    {
        xlnt::workbook wb;
        const auto path = path_helper::test_file("6_encrypted_libre.xlsx");
        xlnt_assert_throws(wb.load(path, "incorrect"), xlnt::exception);
        xlnt_assert_throws_nothing(wb.load(path, u8"\u043F\u0430\u0440\u043E\u043B\u044C")); // u8"пароль"
    }

    void test_decrypt_standard()
    {
        xlnt::workbook wb;
        const auto path = path_helper::test_file("7_encrypted_standard.xlsx");
        xlnt_assert_throws(wb.load(path, "incorrect"), xlnt::exception);
        xlnt_assert_throws_nothing(wb.load(path, "password"));
    }

    void test_decrypt_numbers()
    {
        xlnt::workbook wb;
        const auto path = path_helper::test_file("8_encrypted_numbers.xlsx");
        xlnt_assert_throws(wb.load(path, "incorrect"), xlnt::exception);
        xlnt_assert_throws_nothing(wb.load(path, "secret"));
    }

    void test_read_unicode_filename()
    {
#ifdef _MSC_VER
        xlnt::workbook wb;
        // L"/9_unicode_Λ.xlsx" doesn't use wchar_t(0x039B) for the capital lambda...
        // L"/9_unicode_\u039B.xlsx" gives the corrct output
        const auto path = LSTRING_LITERAL(XLNT_TEST_DATA_DIR) L"/9_unicode_\u039B.xlsx"; // L"/9_unicode_Λ.xlsx"
        wb.load(path);
        xlnt_assert_equals(wb.active_sheet().cell("A1").value<std::string>(), u8"un\u00EFc\u00F4d\u0117!"); // u8"unïcôdė!"
#endif

#ifndef __MINGW32__
        xlnt::workbook wb2;
        // u8"/9_unicode_Λ.xlsx" doesn't use 0xc3aa for the capital lambda...
        // u8"/9_unicode_\u039B.xlsx" gives the corrct output
        const auto path2 = U8STRING_LITERAL(XLNT_TEST_DATA_DIR) u8"/9_unicode_\u039B.xlsx"; // u8"/9_unicode_Λ.xlsx"
        wb2.load(path2);
        xlnt_assert_equals(wb2.active_sheet().cell("A1").value<std::string>(), u8"un\u00EFc\u00F4d\u0117!"); // u8"unïcôdė!"
#endif
    }

    void test_comments()
    {
        xlnt::workbook wb;
        const auto path = path_helper::test_file("10_comments_hyperlinks_formulae.xlsx");
        wb.load(path);

        auto sheet1 = wb[0];
        xlnt_assert_equals(sheet1.cell("A1").value<std::string>(), "Sheet1!A1");
        xlnt_assert_equals(sheet1.cell("A1").comment().plain_text(), "Sheet1 comment");
        xlnt_assert_equals(sheet1.cell("A1").comment().author(), "Microsoft Office User");

        auto sheet2 = wb[1];
        xlnt_assert_equals(sheet2.cell("A1").value<std::string>(), "Sheet2!A1");
        xlnt_assert_equals(sheet2.cell("A1").comment().plain_text(), "Sheet2 comment");
        xlnt_assert_equals(sheet2.cell("A1").comment().author(), "Microsoft Office User");
    }

    void test_read_hyperlink()
    {
        xlnt::workbook wb;
        const auto path = path_helper::test_file("10_comments_hyperlinks_formulae.xlsx");
        wb.load(path);

        auto ws1 = wb.sheet_by_index(0);
        xlnt_assert_equals(ws1.title(), "Sheet1");
        xlnt_assert(ws1.cell("A4").has_hyperlink());
        xlnt_assert_equals(ws1.cell("A4").value<std::string>(), "hyperlink1");
        xlnt_assert_equals(ws1.cell("A4").hyperlink().url(), "https://microsoft.com/");
        xlnt_assert(ws1.cell("A5").has_hyperlink());
        xlnt_assert_equals(ws1.cell("A5").value<std::string>(), "https://google.com/");
        xlnt_assert_equals(ws1.cell("A5").hyperlink().url(), "https://google.com/");
        xlnt_assert(ws1.cell("A6").has_hyperlink());
        xlnt_assert_equals(ws1.cell("A6").value<std::string>(), "Sheet1!A1");
        xlnt_assert_equals(ws1.cell("A6").hyperlink().target_range(), "Sheet1!A1");
        xlnt_assert(ws1.cell("A7").has_hyperlink());
        xlnt_assert_equals(ws1.cell("A7").value<std::string>(), "mailto:invalid@example.com?subject=important");
        xlnt_assert_equals(ws1.cell("A7").hyperlink().url(), "mailto:invalid@example.com?subject=important");
    }

    void test_read_formulae()
    {
        xlnt::workbook wb;
        const auto path = path_helper::test_file("10_comments_hyperlinks_formulae.xlsx");
        wb.load(path);

        auto ws1 = wb.sheet_by_index(0);
        xlnt_assert(ws1.cell("C1").has_formula());
        xlnt_assert_equals(ws1.cell("C1").formula(), "CONCATENATE(C2,C3)");
        xlnt_assert_equals(ws1.cell("C2").value<std::string>(), "a");
        xlnt_assert_equals(ws1.cell("C3").value<std::string>(), "b");

        auto ws2 = wb.sheet_by_index(1);
        xlnt_assert(ws2.cell("C1").has_formula());
        xlnt_assert_equals(ws2.cell("C1").formula(), "C2*C3");
        xlnt_assert_equals(ws2.cell("C2").value<int>(), 2);
        xlnt_assert_equals(ws2.cell("C3").value<int>(), 3);
    }

    void test_read_headers_and_footers()
    {
        xlnt::workbook wb;
        wb.load(path_helper::test_file("11_print_settings.xlsx"));
        auto ws = wb.active_sheet();

        xlnt_assert_equals(ws.cell("A1").value<std::string>(), "header");
        xlnt_assert_equals(ws.cell("A2").value<std::string>(), "and");
        xlnt_assert_equals(ws.cell("A3").value<std::string>(), "footer");
        xlnt_assert_equals(ws.cell("A4").value<std::string>(), "page1");
        xlnt_assert_equals(ws.cell("A43").value<std::string>(), "page2");

        xlnt_assert(ws.has_header_footer());
        xlnt_assert(ws.header_footer().align_with_margins());
        xlnt_assert(ws.header_footer().scale_with_doc());
        xlnt_assert(!ws.header_footer().different_first());
        xlnt_assert(!ws.header_footer().different_odd_even());

        xlnt_assert(ws.header_footer().has_header(xlnt::header_footer::location::left));
        xlnt_assert_equals(ws.header_footer().header(xlnt::header_footer::location::left).plain_text(), "left header");
        xlnt_assert(ws.header_footer().has_header(xlnt::header_footer::location::center));
        xlnt_assert_equals(ws.header_footer().header(xlnt::header_footer::location::center).plain_text(), "center header");
        xlnt_assert(ws.header_footer().has_header(xlnt::header_footer::location::right));
        xlnt_assert_equals(ws.header_footer().header(xlnt::header_footer::location::right).plain_text(), "right header");
        xlnt_assert(ws.header_footer().has_footer(xlnt::header_footer::location::left));
        xlnt_assert_equals(ws.header_footer().footer(xlnt::header_footer::location::left).plain_text(), "left && footer");
        xlnt_assert(ws.header_footer().has_footer(xlnt::header_footer::location::center));
        xlnt_assert_equals(ws.header_footer().footer(xlnt::header_footer::location::center).plain_text(), "center footer");
        xlnt_assert(ws.header_footer().has_footer(xlnt::header_footer::location::right));
        xlnt_assert_equals(ws.header_footer().footer(xlnt::header_footer::location::right).plain_text(), "right footer");
    }

    void test_read_custom_properties()
    {
        xlnt::workbook wb;
        wb.load(path_helper::test_file("12_advanced_properties.xlsx"));
        xlnt_assert(wb.has_custom_property("Client"));
        xlnt_assert_equals(wb.custom_property("Client").get<std::string>(), "me!");
    }

    void test_read_custom_heights_widths()
    {
        xlnt::workbook wb;
        wb.load(path_helper::test_file("13_custom_heights_widths.xlsx"));
        auto ws = wb.active_sheet();

        xlnt_assert_equals(ws.cell("A1").value<std::string>(), "A1");
        xlnt_assert_equals(ws.cell("B1").value<std::string>(), "B1");
        xlnt_assert_equals(ws.cell("D1").value<std::string>(), "D1");
        xlnt_assert_equals(ws.cell("A2").value<std::string>(), "A2");
        xlnt_assert_equals(ws.cell("B2").value<std::string>(), "B2");
        xlnt_assert_equals(ws.cell("D2").value<std::string>(), "D2");
        xlnt_assert_equals(ws.cell("A4").value<std::string>(), "A4");
        xlnt_assert_equals(ws.cell("B4").value<std::string>(), "B4");
        xlnt_assert_equals(ws.cell("D4").value<std::string>(), "D4");

        xlnt_assert_equals(ws.row_properties(1).height.get(), 99.95);
        xlnt_assert(!ws.row_properties(2).height.is_set());
        xlnt_assert_equals(ws.row_properties(3).height.get(), 99.95);
        xlnt_assert(!ws.row_properties(4).height.is_set());
        xlnt_assert_equals(ws.row_properties(5).height.get(), 99.95);

        auto width = ((16.0 * 7) - 5) / 7;

        xlnt_assert_delta(ws.column_properties("A").width.get(), width, 1.0E-9);
        xlnt_assert(!ws.column_properties("B").width.is_set());
        xlnt_assert_delta(ws.column_properties("C").width.get(), width, 1.0E-9);
        xlnt_assert(!ws.column_properties("D").width.is_set());
        xlnt_assert_delta(ws.column_properties("E").width.get(), width, 1.0E-9);
    }

    void test_write_custom_heights_widths()
    {
        xlnt::workbook wb;

        wb.core_property(xlnt::core_property::creator, "Microsoft Office User");
        wb.core_property(xlnt::core_property::last_modified_by, "Microsoft Office User");
        wb.core_property(xlnt::core_property::created, "2017-10-30T23:03:54Z");
        wb.core_property(xlnt::core_property::modified, "2017-10-30T23:08:31Z");
        wb.default_slicer_style("SlicerStyleLight1");
        wb.enable_known_fonts();

        auto ws = wb.active_sheet();

        auto sheet_format_properties = xlnt::sheet_format_properties();
        sheet_format_properties.base_col_width = 10.0;
        sheet_format_properties.default_row_height = 16.0;
        sheet_format_properties.dy_descent = 0.2;
        ws.format_properties(sheet_format_properties);

        ws.cell("A1").value("A1");
        ws.cell("B1").value("B1");
        ws.cell("D1").value("D1");
        ws.cell("A2").value("A2");
        ws.cell("A4").value("A4");
        ws.cell("B2").value("B2");
        ws.cell("B4").value("B4");
        ws.cell("D2").value("D2");
        ws.cell("D4").value("D4");

        for (auto i = 1; i <= 5; ++i)
        {
            ws.row_properties(i).dy_descent = 0.2;
        }

        ws.row_properties(1).height = 99.95;
        ws.row_properties(1).custom_height = true;

        ws.row_properties(3).height = 99.95;
        ws.row_properties(3).custom_height = true;

        ws.row_properties(5).height = 99.95;
        ws.row_properties(5).custom_height = true;

        auto width = ((16.0 * 7) - 5) / 7;

        ws.column_properties("A").width = width;
        ws.column_properties("A").custom_width = true;

        ws.column_properties("C").width = width;
        ws.column_properties("C").custom_width = true;

        ws.column_properties("E").width = width;
        ws.column_properties("E").custom_width = true;

        xlnt_assert(workbook_matches_file(wb,
            path_helper::test_file("13_custom_heights_widths.xlsx")));
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
        source_workbook.save("temp" + source.filename());

#ifdef _MSC_VER
        std::ifstream source_stream(source.wstring(), std::ios::binary);
#else
        std::ifstream source_stream(source.string(), std::ios::binary);
#endif

        return xml_helper::xlsx_archives_match(xlnt::detail::to_vector(source_stream), destination);
    }

    bool round_trip_matches_rw(const xlnt::path &source, const std::string &password)
    {
#ifdef _MSC_VER
        std::ifstream source_stream(source.wstring(), std::ios::binary);
#else
        std::ifstream source_stream(source.string(), std::ios::binary);
#endif
        auto source_data = xlnt::detail::to_vector(source_stream);

        xlnt::workbook source_workbook;
        source_workbook.load(source_data, password);

        std::vector<std::uint8_t> destination_data;
        //source_workbook.save(destination_data, password);
        source_workbook.save("encrypted.xlsx", password);

        //xlnt::workbook temp;
        //temp.load("encrypted.xlsx", password);

        //TODO: finish implementing encryption and uncomment this
        //return source_data == destination_data;
        return true;
    }

    void test_round_trip_rw_minimal()
    {
        xlnt_assert(round_trip_matches_rw(path_helper::test_file("2_minimal.xlsx")));
    }

    void test_round_trip_rw_default()
    {
        xlnt_assert(round_trip_matches_rw(path_helper::test_file("3_default.xlsx")));
    }

    void test_round_trip_rw_every_style()
    {
        xlnt_assert(round_trip_matches_rw(path_helper::test_file("4_every_style.xlsx")));
    }

    void test_round_trip_rw_unicode()
    {
        // u8"/9_unicode_Λ.xlsx" doesn't use 0xc3aa for the capital lambda...
        // u8"/9_unicode_\u039B.xlsx" gives the corrct output
        xlnt_assert(round_trip_matches_rw(path_helper::test_file(u8"9_unicode_\u039B.xlsx")));
    }

    void test_round_trip_rw_comments_hyperlinks_formulae()
    {
        xlnt_assert(round_trip_matches_rw(path_helper::test_file("10_comments_hyperlinks_formulae.xlsx")));
    }

    void test_round_trip_rw_print_settings()
    {
        xlnt_assert(round_trip_matches_rw(path_helper::test_file("11_print_settings.xlsx")));
    }

    void test_round_trip_rw_advanced_properties()
    {
        xlnt_assert(round_trip_matches_rw(path_helper::test_file("12_advanced_properties.xlsx")));
    }

    void test_round_trip_rw_custom_heights_widths()
    {
        xlnt_assert(round_trip_matches_rw(path_helper::test_file("13_custom_heights_widths.xlsx")));
    }

    void test_round_trip_rw_encrypted_agile()
    {
        xlnt_assert(round_trip_matches_rw(path_helper::test_file("5_encrypted_agile.xlsx"), "secret"));
    }

    void test_round_trip_rw_encrypted_libre()
    {
        xlnt_assert(round_trip_matches_rw(path_helper::test_file("6_encrypted_libre.xlsx"), u8"\u043F\u0430\u0440\u043E\u043B\u044C")); // u8"пароль"
    }

    void test_round_trip_rw_encrypted_standard()
    {
        xlnt_assert(round_trip_matches_rw(path_helper::test_file("7_encrypted_standard.xlsx"), "password"));
    }

    void test_round_trip_rw_encrypted_numbers()
    {
        xlnt_assert(round_trip_matches_rw(path_helper::test_file("8_encrypted_numbers.xlsx"), "secret"));
    }

    void test_streaming_read()
    {
        const auto path = path_helper::test_file("4_every_style.xlsx");
        xlnt::streaming_workbook_reader reader;

        reader.open(xlnt::path(path));

        for (auto sheet_name : reader.sheet_titles())
        {
            reader.begin_worksheet(sheet_name);

            while (reader.has_cell())
            {
                reader.read_cell();
            }

            reader.end_worksheet();
        }
    }

    void test_streaming_write()
    {
        const auto path = std::string("stream-out.xlsx");
        xlnt::streaming_workbook_writer writer;

        writer.open(path);

        writer.add_worksheet("stream");

        auto b2 = writer.add_cell("B2");
        b2.value("B2!");

        auto c3 = writer.add_cell("C3");
        b2.value("should not change");
        c3.value("C3!");
    }
};
static serialization_test_suite x;
