#pragma once

#include <fstream>
#include <iostream>
#include <cxxtest/TestSuite.h>

#include <helpers/path_helper.hpp>
#include <xlnt/workbook/workbook.hpp>

class test_consume_xlsx : public CxxTest::TestSuite
{
public:
    void test_decrypt_agile()
    {
        xlnt::workbook wb;
        wb.load(path_helper::get_data_directory("11_encrypted_excel_2016.xlsx"), "secret");
    }

    void test_decrypt_libre_office()
    {
        xlnt::workbook wb;
        wb.load(path_helper::get_data_directory("12_encrypted_libre_office.xlsx"), "secret");
    }

    void test_decrypt_standard()
    {
        xlnt::workbook wb;
        wb.load(path_helper::get_data_directory("13_encrypted_excel_2007.xlsx"), "password");
    }

    void test_decrypt_numbers()
    {
        xlnt::workbook wb;
        wb.load(path_helper::get_data_directory("14_encrypted_numbers.xlsx"), "secret");
    }

    void test_comments()
    {
        xlnt::workbook wb;
        wb.load("data/15_basic_comments.xlsx");

        auto sheet1 = wb[0];
        TS_ASSERT_EQUALS(sheet1.cell("A1").value<std::string>(), "Sheet1!A1");
        TS_ASSERT_EQUALS(sheet1.cell("A1").comment().plain_text(), "Sheet1 comment");
        TS_ASSERT_EQUALS(sheet1.cell("A1").comment().author(), "Microsoft Office User");

        auto sheet2 = wb[1];
        TS_ASSERT_EQUALS(sheet2.cell("A1").value<std::string>(), "Sheet2!A1");
        TS_ASSERT_EQUALS(sheet2.cell("A1").comment().plain_text(), "Sheet2 comment");
        TS_ASSERT_EQUALS(sheet2.cell("A1").comment().author(), "Microsoft Office User");
    }

    void test_read_unicode_filename()
    {
#ifdef _MSC_VER
        xlnt::workbook wb;
        wb.load(L"data\\16_unicode_Λ.xlsx");
        TS_ASSERT_EQUALS(wb.active_sheet().cell("A1").value<std::string>(), "unicode!");
#endif
#ifndef __MINGW32__
        xlnt::workbook wb2;
        wb2.load(u8"data/16_unicode_Λ.xlsx");
        TS_ASSERT_EQUALS(wb2.active_sheet().cell("A1").value<std::string>(), "unicode!");
#endif
    }

    void test_read_hyperlink()
    {
        xlnt::workbook wb;
        wb.load("data/17_with_hyperlink.xlsx");
        TS_ASSERT(wb.active_sheet().cell("A1").has_hyperlink());
        TS_ASSERT_EQUALS(wb.active_sheet().cell("A1").hyperlink(),
            "https://fr.wikipedia.org/wiki/Ille-et-Vilaine");
    }

    void test_read_headers_and_footers()
    {
        xlnt::workbook wb;
        wb.load("data/18_headers_and_footers.xlsx");
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
        wb.load("data/21_custom_properties.xlsx");
        TS_ASSERT(wb.has_custom_property("Client"));
        TS_ASSERT_EQUALS(wb.custom_property("Client"), "me!");
    }
};
