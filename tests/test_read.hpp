#pragma once

#include <fstream>
#include <iostream>

#include <cxxtest/TestSuite.h>

#include <xlnt/s11n/excel_serializer.hpp>
#include <xlnt/s11n/manifest_serializer.hpp>
#include <xlnt/s11n/workbook_serializer.hpp>
#include <xlnt/s11n/xml_serializer.hpp>
#include <xlnt/workbook/manifest.hpp>

#include "helpers/path_helper.hpp"

class test_read : public CxxTest::TestSuite
{
public:
    
    xlnt::workbook standard_workbook()
    {
        xlnt::workbook wb;
        auto path = PathHelper::GetDataDirectory("/genuine/empty.xlsx");
        wb.load(path);
        
        return wb;
    }

    void test_read_standard_workbook()
    {
        TS_ASSERT_DIFFERS(standard_workbook(), nullptr);
    }

    void test_read_standard_workbook_from_fileobj()
    {
        auto path = PathHelper::GetDataDirectory("/genuine/empty.xlsx");
        std::ifstream fo(path, std::ios::binary);
        
        xlnt::workbook wb;
        xlnt::excel_serializer serializer(wb);
        
        serializer.load_stream_workbook(fo);
        
        TS_ASSERT_DIFFERS(wb, nullptr);
    }

    void test_read_worksheet()
    {
        auto wb = standard_workbook();
        auto sheet2 = wb.get_sheet_by_name("Sheet2 - Numbers");
        
        TS_ASSERT_DIFFERS(sheet2, nullptr);
        TS_ASSERT_EQUALS("This is cell G5", sheet2.get_cell("G5").get_value<std::string>());
        TS_ASSERT_EQUALS(18, sheet2.get_cell("D18").get_value<int>());
        TS_ASSERT_EQUALS(true, sheet2.get_cell("G9").get_value<bool>());
        TS_ASSERT_EQUALS(false, sheet2.get_cell("G10").get_value<bool>());
    }

    void test_read_nostring_workbook()
    {
        auto path = PathHelper::GetDataDirectory("/genuine/empty-no-string.xlsx");
        
        xlnt::workbook wb;
        xlnt::excel_serializer serializer(wb);
        
        serializer.load_workbook(path);
        
        TS_ASSERT_DIFFERS(wb, nullptr);
    }

    void test_read_empty_file()
    {
        auto path = PathHelper::GetDataDirectory("/reader/null_file.xlsx");
        
        xlnt::workbook wb;
        xlnt::excel_serializer serializer(wb);
        
        TS_ASSERT_THROWS(serializer.load_workbook(path), xlnt::invalid_file_exception);
    }

    void test_read_empty_archive()
    {
        auto path = PathHelper::GetDataDirectory("/reader/null_archive.xlsx");
        
        xlnt::workbook wb;
        xlnt::excel_serializer serializer(wb);
        
        TS_ASSERT_THROWS(serializer.load_workbook(path), xlnt::invalid_file_exception);
    }
    
    void test_read_workbook_with_no_properties()
    {
        auto path = PathHelper::GetDataDirectory("/genuine/empty_with_no_properties.xlsx");

        xlnt::workbook wb;
        xlnt::excel_serializer serializer(wb);
        
        serializer.load_workbook(path);
    }

    xlnt::workbook workbook_with_styles()
    {
        auto path = PathHelper::GetDataDirectory("/genuine/empty-with-styles.xlsx");

        xlnt::workbook wb;
        xlnt::excel_serializer serializer(wb);
        
        serializer.load_workbook(path);
        
        return wb;
    }

    void test_read_workbook_with_styles_general()
    {
        auto wb = workbook_with_styles();
        auto ws = wb["Sheet1"];
        auto format = ws.get_cell("A1").get_number_format();
        auto expected = xlnt::number_format::general();
        TS_ASSERT_EQUALS(format, expected);
    }
    
    void test_read_workbook_with_styles_date()
    {
        auto wb = workbook_with_styles();
        auto ws = wb["Sheet1"];
        auto format = ws.get_cell("A2").get_number_format();
        auto expected = xlnt::number_format::date_xlsx14();
        TS_ASSERT_EQUALS(format, expected);
    }
    
    void test_read_workbook_with_styles_number()
    {
        auto wb = workbook_with_styles();
        auto ws = wb["Sheet1"];
        auto code = ws.get_cell("A3").get_number_format();
        auto expected = xlnt::number_format::number_00();
        TS_ASSERT_EQUALS(code, expected);
    }
    
    void test_read_workbook_with_styles_time()
    {
        auto wb = workbook_with_styles();
        auto ws = wb["Sheet1"];
        auto code = ws.get_cell("A4").get_number_format();
        auto expected = xlnt::number_format::date_time3();
        TS_ASSERT_EQUALS(code, expected);
    }
    
    void test_read_workbook_with_styles_percentage()
    {
        auto wb = workbook_with_styles();
        auto ws = wb["Sheet1"];
        auto code = ws.get_cell("A5").get_number_format();
        auto expected = xlnt::number_format::percentage_00();
        TS_ASSERT_EQUALS(code, expected);
    }
    
    xlnt::workbook date_mac_1904()
    {
        auto path = PathHelper::GetDataDirectory("/reader/date_1904.xlsx");
        
        xlnt::workbook wb;
        xlnt::excel_serializer serializer(wb);
        
        serializer.load_workbook(path);
        
        return wb;
    }
    
    xlnt::workbook date_std_1900()
    {
        auto path = PathHelper::GetDataDirectory("/reader/date_1900.xlsx");

        xlnt::workbook wb;
        xlnt::excel_serializer serializer(wb);
        
        serializer.load_workbook(path);
        
        return wb;
    }
    
    void test_read_win_base_date()
    {
        auto wb = date_std_1900();
        TS_ASSERT_EQUALS(wb.get_properties().excel_base_date, xlnt::calendar::windows_1900);
    }
    
    void test_read_mac_base_date()
    {
        auto wb = date_mac_1904();
        TS_ASSERT_EQUALS(wb.get_properties().excel_base_date, xlnt::calendar::mac_1904);
    }

    void test_read_date_style_win()
    {
        auto wb = date_std_1900();
        auto ws = wb["Sheet1"];
        TS_ASSERT_EQUALS(ws.get_cell("A1").get_number_format(), xlnt::number_format::date_xlsx14());
    }

    void test_read_date_style_mac()
    {
        auto wb = date_mac_1904();
        auto ws = wb["Sheet1"];
        TS_ASSERT_EQUALS(ws.get_cell("A1").get_number_format(), xlnt::number_format::date_xlsx14());
    }

    void test_read_compare_mac_win_dates()
    {
        auto wb_mac = date_mac_1904();
        auto ws_mac = wb_mac["Sheet1"];
        auto wb_win = date_std_1900();
        auto ws_win = wb_win["Sheet1"];
        xlnt::datetime dt(2011, 10, 31);
        TS_ASSERT_EQUALS(ws_mac.get_cell("A1").get_value<xlnt::datetime>(), dt);
        TS_ASSERT_EQUALS(ws_win.get_cell("A1").get_value<xlnt::datetime>(), dt);
        TS_ASSERT_EQUALS(ws_mac.get_cell("A1").get_value<xlnt::datetime>(), ws_win.get_cell("A1").get_value<xlnt::datetime>());
    }
    
    void test_repair_central_directory()
    {
        std::string data_a = "foobarbaz" + xlnt::excel_serializer::central_directory_signature();
        std::string data_b = "bazbarfoo12345678901234567890";
        
        auto f = xlnt::excel_serializer::repair_central_directory(data_a + data_b);
        TS_ASSERT_EQUALS(f, data_a + data_b.substr(0, 18));
        
        f = xlnt::excel_serializer::repair_central_directory(data_b);
        TS_ASSERT_EQUALS(f, data_b);
    }
    
    void test_read_no_theme()
    {
        auto path = PathHelper::GetDataDirectory("/genuine/libreoffice_nrt.xlsx");
        
        xlnt::workbook wb;
        xlnt::excel_serializer serializer(wb);
        
        serializer.load_workbook(path);
        
        TS_ASSERT_DIFFERS(wb, nullptr);
    }
    
    void _test_read_complex_formulae()
    {
        /*
        auto path = PathHelper::GetDataDirectory("/reader/formulae.xlsx");
        auto wb = xlnt::reader::load_workbook(path);
        auto ws = wb.get_active_sheet();

        // Test normal forumlae
        TS_ASSERT(!ws.get_cell("A1").has_formula());
        TS_ASSERT(!ws.get_cell("A2").has_formula());
        TS_ASSERT(ws.get_cell("A3").has_formula());
        TS_ASSERT(ws.get_formula_attributes().find("A3") == ws.get_formula_attributes().end());
        TS_ASSERT(ws.get_cell("A3").get_formula() == "12345");
        TS_ASSERT(ws.get_cell("A4").has_formula());
        TS_ASSERT(ws.get_formula_attributes().find("A3") == ws.get_formula_attributes().end());
        ws.get_cell("A4").set_formula("A2+A3");
        TS_ASSERT(ws.get_cell("A5").has_formula());
        TS_ASSERT(ws.get_formula_attributes().find("A5") == ws.get_formula_attributes().end());
        ws.get_cell("A5").set_formula("SUM(A2:A4)");

        // Test unicode
        std::string expected = "=IF(ISBLANK(B16), \"D\xFCsseldorf\", B16)";
        TS_ASSERT(ws.get_cell("A16").get_formula() == expected);

        // Test shared forumlae
        TS_ASSERT(ws.get_cell("B7").get_data_type() == "f");
        TS_ASSERT(ws.formula_attributes["B7"]["t"] == "shared");
        TS_ASSERT(ws.formula_attributes["B7"]["si"] == "0");
        TS_ASSERT(ws.formula_attributes["B7"]["ref"] == "B7:E7");
        TS_ASSERT(ws.get_cell("B7").value == "=B4*2");
        TS_ASSERT(ws.get_cell("C7").get_data_type() == "f");
        TS_ASSERT(ws.formula_attributes["C7"]["t"] == "shared");
        TS_ASSERT(ws.formula_attributes["C7"]["si"] == "0");
        TS_ASSERT("ref" not in ws.formula_attributes["C7"]);
        TS_ASSERT(ws.get_cell("C7").value == "=");
        TS_ASSERT(ws.get_cell("D7").get_data_type() == "f");
        TS_ASSERT(ws.formula_attributes["D7"]["t"] == "shared");
        TS_ASSERT(ws.formula_attributes["D7"]["si"] == "0");
        TS_ASSERT("ref" not in ws.formula_attributes["D7"]);
        TS_ASSERT(ws.get_cell("D7").value == "=");
        TS_ASSERT(ws.get_cell("E7").get_data_type() == "f");
        TS_ASSERT(ws.formula_attributes["E7"]["t"] == "shared");
        TS_ASSERT(ws.formula_attributes["E7"]["si"] == "0");
        TS_ASSERT("ref" not in ws.formula_attributes["E7"]);
        TS_ASSERT(ws.get_cell("E7").value == "=");

        // Test array forumlae
        TS_ASSERT(ws.get_cell("C10").get_data_type() == "f");
        TS_ASSERT("ref" not in ws.formula_attributes["C10"]["ref"]);
        TS_ASSERT(ws.formula_attributes["C10"]["t"] == "array");
        TS_ASSERT("si" not in ws.formula_attributes["C10"]);
        TS_ASSERT(ws.formula_attributes["C10"]["ref"] == "C10:C14");
        TS_ASSERT(ws.get_cell("C10").value == "=SUM(A10:A14*B10:B14)");
        TS_ASSERT(ws.get_cell("C11").get_data_type() != "f");
        */
    }
    
    void test_data_only()
    {
        auto path = PathHelper::GetDataDirectory("/reader/formulae.xlsx");
        
        
        xlnt::workbook wb;
        xlnt::excel_serializer serializer(wb);
        
        serializer.load_workbook(path, false, true);

        auto ws = wb.get_active_sheet();

        TS_ASSERT(ws.get_formula_attributes().empty());
        TS_ASSERT(ws.get_parent().get_data_only());
        TS_ASSERT(ws.get_cell("A2").get_data_type() == xlnt::cell::type::numeric);
        TS_ASSERT(ws.get_cell("A2").get_value<int>() == 12345);
        TS_ASSERT(!ws.get_cell("A2").has_formula());
        TS_ASSERT(ws.get_cell("A3").get_data_type() == xlnt::cell::type::numeric);
        TS_ASSERT(ws.get_cell("A3").get_value<int>() == 12345);
        TS_ASSERT(!ws.get_cell("A3").has_formula());
        TS_ASSERT(ws.get_cell("A4").get_data_type() == xlnt::cell::type::numeric);
        TS_ASSERT(ws.get_cell("A4").get_value<int>() == 24690);
        TS_ASSERT(!ws.get_cell("A4").has_formula());
        TS_ASSERT(ws.get_cell("A5").get_data_type() == xlnt::cell::type::numeric);
        TS_ASSERT(ws.get_cell("A5").get_value<int>() == 49380);
        TS_ASSERT(!ws.get_cell("A5").has_formula());
    }
    
    void test_detect_worksheets()
    {
        {
            auto path = PathHelper::GetDataDirectory("/reader/bug137.xlsx");
            xlnt::zip_file archive(path);
            std::vector<std::pair<std::string, std::string>> expected =
            {
                {"xl/worksheets/sheet1.xml", "Sheet1"}
            };
            TS_ASSERT_EQUALS(xlnt::detect_worksheets(archive), expected);
        }

        {
            auto path = PathHelper::GetDataDirectory("/reader/contains_chartsheets.xlsx");
            xlnt::zip_file archive(path);

            std::vector<std::pair<std::string, std::string>> expected =
            {
                {"xl/worksheets/sheet1.xml", "data"},
                {"xl/worksheets/sheet2.xml", "moredata"}
            };

            TS_ASSERT_EQUALS(xlnt::detect_worksheets(archive), expected);
        }

        {
            auto path = PathHelper::GetDataDirectory("/reader/bug304.xlsx");
            xlnt::zip_file archive(path);

            std::vector<std::pair<std::string, std::string>> expected =
            {
                {"xl/worksheets/sheet3.xml", "Sheet1"},
                {"xl/worksheets/sheet2.xml", "Sheet2"},
                {"xl/worksheets/sheet.xml", "Sheet3"}
            };

            TS_ASSERT_EQUALS(xlnt::detect_worksheets(archive), expected);
        }
    }
    
    void test_read_rels()
    {
        {
            std::vector<xlnt::relationship> expected =
            {
                {xlnt::relationship::type::theme, "rId3", "xl/theme/theme1.xml"},
                {xlnt::relationship::type::worksheet, "rId2", "xl/worksheets/sheet1.xml"},
                {xlnt::relationship::type::chartsheet, "rId1", "xl/chartsheets/sheet1.xml"},
                {xlnt::relationship::type::shared_strings, "rId5", "xl/sharedStrings.xml"},
                {xlnt::relationship::type::styles, "rId4", "xl/styles.xml"}
            };

            auto path = PathHelper::GetDataDirectory("/reader/bug137.xlsx");
            xlnt::zip_file archive(path);

            TS_ASSERT_EQUALS(xlnt::read_relationships(archive, "xl/workbook.xml"), expected);
        }

        {
            std::vector<xlnt::relationship> expected =
            {
                {xlnt::relationship::type::custom_xml, "rId8", "../customXml/item3.xml"},
                {xlnt::relationship::type::worksheet, "rId3", "xl/worksheets/sheet.xml"},
                {xlnt::relationship::type::custom_xml, "rId7", "../customXml/item2.xml"},
                {xlnt::relationship::type::worksheet, "rId2", "xl/worksheets/sheet2.xml"},
                {xlnt::relationship::type::worksheet, "rId1", "xl/worksheets/sheet3.xml"},
                {xlnt::relationship::type::custom_xml, "rId6", "../customXml/item1.xml"},
                {xlnt::relationship::type::styles, "rId5", "xl/styles.xml"},
                {xlnt::relationship::type::theme, "rId4", "xl/theme/theme.xml"}
            };

            auto path = PathHelper::GetDataDirectory("/reader/bug304.xlsx");
            xlnt::zip_file archive(path);

            TS_ASSERT_EQUALS(xlnt::read_relationships(archive, "xl/workbook.xml"), expected);
        }
    }
    
    void test_read_content_types()
    {
        std::vector<std::pair<std::string, std::string>> expected = 
        {
            {"/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"},
            {"/xl/worksheets/sheet1.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"},
            {"/xl/chartsheets/sheet1.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"},
            {"/xl/worksheets/sheet2.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"},
            {"/xl/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml"},
            {"/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"},
            {"/xl/sharedStrings.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"},
            {"/xl/drawings/drawing1.xml", "application/vnd.openxmlformats-officedocument.drawing+xml"},
            {"/xl/charts/chart1.xml", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"},
            {"/xl/drawings/drawing2.xml", "application/vnd.openxmlformats-officedocument.drawing+xml"},
            {"/xl/charts/chart2.xml", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"},
            {"/xl/calcChain.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"},
            {"/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml"},
            {"/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml"}
        };
        
        auto path = PathHelper::GetDataDirectory("/reader/contains_chartsheets.xlsx");

        xlnt::workbook wb;
        xlnt::manifest_serializer serializer(wb.get_manifest());
        
        xlnt::zip_file archive;
        archive.load(path);
        
        serializer.read_manifest(xlnt::xml_serializer::deserialize(archive.read("[Content Types].xml")));
        
        auto &result = wb.get_manifest().get_override_types();

        if(result.size() != expected.size())
        {
            TS_ASSERT_EQUALS(result.size(), expected.size());
            return;
        }

        for(std::size_t i = 0; i < expected.size(); i++)
        {
            TS_ASSERT_EQUALS(result[i].get_part_name(), expected[i].first);
            TS_ASSERT_EQUALS(result[i].get_content_type(), expected[i].second);
        }
    }
    
    void test_read_sheets()
    {
        {
            auto path = PathHelper::GetDataDirectory("/reader/bug137.xlsx");

            xlnt::workbook wb;
            xlnt::workbook_serializer serializer(wb);
            
            auto sheets = serializer.read_sheets();
            std::vector<std::pair<std::string, std::string>> expected =
            {{"rId1", "Chart1"}, {"rId2", "Sheet1"}};
            TS_ASSERT_EQUALS(sheets, expected);
        }
        
        {
            auto path = PathHelper::GetDataDirectory("/reader/bug304.xlsx");

            xlnt::workbook wb;
            xlnt::workbook_serializer serializer(wb);
            
            auto sheets = serializer.read_sheets();
            std::vector<std::pair<std::string, std::string>> expected =
            {{"rId1", "Sheet1"}, {"rId2", "Sheet2"}, {"rId3", "Sheet3"}};
            TS_ASSERT_EQUALS(sheets, expected);
        }
    }
    
    void test_guess_types()
    {
        bool guess;
        xlnt::cell::type dtype;
        std::vector<std::pair<bool, xlnt::cell::type>> test_cases = {{true, xlnt::cell::type::numeric}, {false, xlnt::cell::type::string}};
        
        for(const auto &expected : test_cases)
        {
            std::tie(guess, dtype) = expected;
            auto path = PathHelper::GetDataDirectory("/genuine/guess_types.xlsx");
            
            xlnt::workbook wb;
            xlnt::excel_serializer serializer(wb);
            
            serializer.load_workbook(path, guess);

            auto ws = wb.get_active_sheet();
            TS_ASSERT(ws.get_cell("D2").get_data_type() == dtype);
        }
    }
    
    void test_read_autofilter()
    {
        auto path = PathHelper::GetDataDirectory("/reader/bug275.xlsx");

        xlnt::workbook wb;
        xlnt::excel_serializer serializer(wb);
        
        serializer.load_workbook(path);
        
        auto ws = wb.get_active_sheet();
        TS_ASSERT_EQUALS(ws.get_auto_filter().to_string(), "A1:B6");
    }
    
    void test_bad_formats_xlsb()
    {
        auto path = PathHelper::GetDataDirectory("/genuine/a.xlsb");
        
        xlnt::workbook wb;
        xlnt::excel_serializer serializer(wb);
        
        TS_ASSERT_THROWS(serializer.load_workbook(path), xlnt::invalid_file_exception);
    }
    
    void test_bad_formats_xls()
    {
        auto path = PathHelper::GetDataDirectory("/genuine/a.xls");
        
        xlnt::workbook wb;
        xlnt::excel_serializer serializer(wb);
        
        TS_ASSERT_THROWS(serializer.load_workbook(path), xlnt::invalid_file_exception);
    }
    
    void test_bad_formats_no()
    {
        auto path = PathHelper::GetDataDirectory("/genuine/a.no-format");
        
        xlnt::workbook wb;
        xlnt::excel_serializer serializer(wb);
        
        TS_ASSERT_THROWS(serializer.load_workbook(path), xlnt::invalid_file_exception);
    }
};
