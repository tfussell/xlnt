#pragma once

#include <fstream>
#include <iostream>
#include <cxxtest/TestSuite.h>

#include <helpers/path_helper.hpp>
#include <xlnt/workbook/workbook.hpp>

class test_consume_xlsx : public CxxTest::TestSuite
{
public:
    void test_non_xlsx()
    {
        xlnt::workbook wb;
        const auto path = path_helper::data_directory("1_powerpoint_presentation.xlsx");
        TS_ASSERT_THROWS(wb.load(path), xlnt::invalid_file);
    }

    void test_decrypt_agile()
    {
        xlnt::workbook wb;
        const auto path = path_helper::data_directory("5_encrypted_agile.xlsx");
        TS_ASSERT_THROWS_NOTHING(wb.load(path, "secret"));
    }

    void test_decrypt_libre_office()
    {
        xlnt::workbook wb;
        const auto path = path_helper::data_directory("6_encrypted_libre.xlsx");
        TS_ASSERT_THROWS_NOTHING(wb.load(path, "secret"));
    }

    void test_decrypt_standard()
    {
        xlnt::workbook wb;
        const auto path = path_helper::data_directory("7_encrypted_standard.xlsx");
        TS_ASSERT_THROWS_NOTHING(wb.load(path, "password"));
    }

    void test_decrypt_numbers()
    {
        xlnt::workbook wb;
        const auto path = path_helper::data_directory("8_encrypted_numbers.xlsx");
        TS_ASSERT_THROWS_NOTHING(wb.load(path, "secret"));
    }

    void test_read_unicode_filename()
    {
#ifdef _MSC_VER
        xlnt::workbook wb;
        const auto path = LSTRING_LITERAL(XLNT_TEST_DATA_DIR) L"/9_unicode_filename_Λ.xlsx";
        wb.load(path);
        TS_ASSERT_EQUALS(wb.active_sheet().cell("A1").value<std::string>(), "unicode!");
#endif

#ifndef __MINGW32__
        xlnt::workbook wb2;
        const auto path = U8STRING_LITERAL(XLNT_TEST_DATA_DIR) u8"/9_unicode_filename_Λ.xlsx";
        wb2.load(path);
        TS_ASSERT_EQUALS(wb2.active_sheet().cell("A1").value<std::string>(), "unicode!");
#endif
    }

    void test_comments()
    {
        xlnt::workbook wb;
        const auto path = path_helper::data_directory("10_comments_hyperlinks_formulae.xlsx");
        wb.load(path);

        auto sheet1 = wb[0];
        TS_ASSERT_EQUALS(sheet1.cell("A1").value<std::string>(), "Sheet1!A1");
        TS_ASSERT_EQUALS(sheet1.cell("A1").comment().plain_text(), "Sheet1 comment");
        TS_ASSERT_EQUALS(sheet1.cell("A1").comment().author(), "Microsoft Office User");

        auto sheet2 = wb[1];
        TS_ASSERT_EQUALS(sheet2.cell("A1").value<std::string>(), "Sheet2!A1");
        TS_ASSERT_EQUALS(sheet2.cell("A1").comment().plain_text(), "Sheet2 comment");
        TS_ASSERT_EQUALS(sheet2.cell("A1").comment().author(), "Microsoft Office User");
    }

    void test_read_hyperlink()
    {
        xlnt::workbook wb;
        const auto path = path_helper::data_directory("10_comments_hyperlinks_formulae.xlsx");
        wb.load(path);

        auto ws1 = wb.sheet_by_index(0);
        TS_ASSERT_EQUALS(ws1.title(), "Sheet1");
        TS_ASSERT(ws1.cell("A4").has_hyperlink());
        TS_ASSERT_EQUALS(ws1.cell("A4").value<std::string>(), "hyperlink1");
        TS_ASSERT_EQUALS(ws1.cell("A4").hyperlink(), "https://microsoft.com/");
        TS_ASSERT(ws1.cell("A5").has_hyperlink());
        TS_ASSERT_EQUALS(ws1.cell("A5").value<std::string>(), "https://google.com/");
        TS_ASSERT_EQUALS(ws1.cell("A5").hyperlink(), "https://google.com/");
        //TS_ASSERT(ws1.cell("A6").has_hyperlink());
        TS_ASSERT_EQUALS(ws1.cell("A6").value<std::string>(), "Sheet1!A1");
        //TS_ASSERT_EQUALS(ws1.cell("A6").hyperlink(), "Sheet1!A1");
        TS_ASSERT(ws1.cell("A7").has_hyperlink());
        TS_ASSERT_EQUALS(ws1.cell("A7").value<std::string>(), "mailto:invalid@example.com?subject=important");
        TS_ASSERT_EQUALS(ws1.cell("A7").hyperlink(), "mailto:invalid@example.com?subject=important");

    }

    void test_read_formulae()
    {
        xlnt::workbook wb;
        const auto path = path_helper::data_directory("10_comments_hyperlinks_formulae.xlsx");
        wb.load(path);

        auto ws1 = wb.sheet_by_index(0);
        TS_ASSERT_EQUALS(ws1.cell("C1").value<std::string>(), "ab");
        TS_ASSERT(ws1.cell("C1").has_formula());
        TS_ASSERT_EQUALS(ws1.cell("C1").formula(), "CONCATENATE(C2,C3)");
        TS_ASSERT_EQUALS(ws1.cell("C2").value<std::string>(), "a");
        TS_ASSERT_EQUALS(ws1.cell("C3").value<std::string>(), "b");

        auto ws2 = wb.sheet_by_index(1);
        TS_ASSERT_EQUALS(ws2.cell("C1").value<int>(), 6);
        TS_ASSERT(ws2.cell("C1").has_formula());
        TS_ASSERT_EQUALS(ws2.cell("C1").formula(), "C2*C3");
        TS_ASSERT_EQUALS(ws2.cell("C2").value<int>(), 2);
        TS_ASSERT_EQUALS(ws2.cell("C3").value<int>(), 3);
    }

    void test_read_headers_and_footers()
    {
        xlnt::workbook wb;
        wb.load(path_helper::data_directory("11_print_settings.xlsx"));
        auto ws = wb.active_sheet();

        TS_ASSERT_EQUALS(ws.cell("A1").value<std::string>(), "header");
        TS_ASSERT_EQUALS(ws.cell("A2").value<std::string>(), "and");
        TS_ASSERT_EQUALS(ws.cell("A3").value<std::string>(), "footer");
        TS_ASSERT_EQUALS(ws.cell("A4").value<std::string>(), "page1");
        TS_ASSERT_EQUALS(ws.cell("A43").value<std::string>(), "page2");

        TS_ASSERT(ws.has_header_footer());
        TS_ASSERT(ws.header_footer().align_with_margins());
        TS_ASSERT(ws.header_footer().scale_with_doc());
        TS_ASSERT(!ws.header_footer().different_first());
        TS_ASSERT(!ws.header_footer().different_odd_even());

        TS_ASSERT(ws.header_footer().has_header(xlnt::header_footer::location::left));
        TS_ASSERT_EQUALS(ws.header_footer().header(xlnt::header_footer::location::left).plain_text(), "left header");
        TS_ASSERT(ws.header_footer().has_header(xlnt::header_footer::location::center));
        TS_ASSERT_EQUALS(ws.header_footer().header(xlnt::header_footer::location::center).plain_text(), "center header");
        TS_ASSERT(ws.header_footer().has_header(xlnt::header_footer::location::right));
        TS_ASSERT_EQUALS(ws.header_footer().header(xlnt::header_footer::location::right).plain_text(), "right header");
        TS_ASSERT(ws.header_footer().has_footer(xlnt::header_footer::location::left));
        TS_ASSERT_EQUALS(ws.header_footer().footer(xlnt::header_footer::location::left).plain_text(), "left && footer");
        TS_ASSERT(ws.header_footer().has_footer(xlnt::header_footer::location::center));
        TS_ASSERT_EQUALS(ws.header_footer().footer(xlnt::header_footer::location::center).plain_text(), "center footer");
        TS_ASSERT(ws.header_footer().has_footer(xlnt::header_footer::location::right));
        TS_ASSERT_EQUALS(ws.header_footer().footer(xlnt::header_footer::location::right).plain_text(), "right footer");
    }

    void test_read_custom_properties()
    {
        xlnt::workbook wb;
        wb.load(path_helper::data_directory("12_advanced_properties.xlsx"));
        TS_ASSERT(wb.has_custom_property("Client"));
        TS_ASSERT_EQUALS(wb.custom_property("Client").get<std::string>(), "me!");
    }
};
