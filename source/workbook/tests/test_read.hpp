#pragma once

#include <fstream>
#include <iostream>
#include <cxxtest/TestSuite.h>

#include <helpers/path_helper.hpp>
#include <xlnt/cell/text.hpp>
#include <xlnt/cell/text_run.hpp>
#include <xlnt/packaging/manifest.hpp>

class test_read : public CxxTest::TestSuite
{
public:
    
    xlnt::workbook standard_workbook()
    {
        xlnt::workbook wb;
        auto path = path_helper::get_data_directory("genuine/empty.xlsx");
        wb.load(path);
        
        return wb;
    }

    void test_read_standard_workbook()
    {
        TS_ASSERT_THROWS_NOTHING(standard_workbook());
    }

    void test_read_standard_workbook_from_fileobj()
    {
        auto path = path_helper::get_data_directory("genuine/empty.xlsx");
        std::ifstream fo(path.to_string(), std::ios::binary);

        xlnt::workbook wb;
		wb.load(fo);

		TS_ASSERT(wb.get_sheet_titles().size() == 1);
    }

    void test_read_worksheet()
    {
        auto wb = standard_workbook();
        auto sheet2 = wb.get_sheet_by_title("Sheet2 - Numbers");
        
        TS_ASSERT_DIFFERS(sheet2, nullptr);
        TS_ASSERT_EQUALS("This is cell G5", sheet2.get_cell("G5").get_value<std::string>());
        TS_ASSERT_EQUALS(18, sheet2.get_cell("D18").get_value<int>());
        TS_ASSERT_EQUALS(true, sheet2.get_cell("G9").get_value<bool>());
        TS_ASSERT_EQUALS(false, sheet2.get_cell("G10").get_value<bool>());
    }

    void test_read_nostring_workbook()
    {
        auto path = path_helper::get_data_directory("genuine/empty-no-string.xlsx");
        
        xlnt::workbook wb;
        TS_ASSERT_THROWS_NOTHING(wb.load(path));
    }

    void test_read_empty_file()
    {
        xlnt::workbook wb;
        TS_ASSERT_THROWS(wb.load(path_helper::get_data_directory("reader/null_file.xlsx")), xlnt::invalid_file);
    }

    void test_read_empty_archive()
    {
		xlnt::workbook wb;
		TS_ASSERT_THROWS(wb.load(path_helper::get_data_directory("reader/null_archive.xlsx")), xlnt::invalid_file);
    }
    
    void test_read_workbook_with_no_properties()
    {
		xlnt::workbook wb;
		TS_ASSERT_THROWS_NOTHING(wb.load(path_helper::get_data_directory("genuine/empty_with_no_properties.xlsx")));
    }

    xlnt::workbook workbook_with_styles()
    {
		xlnt::workbook wb;
		wb.load(path_helper::get_data_directory("genuine/empty-with-styles.xlsx"));

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
    
    void test_read_charset_excel()
    {
        xlnt::workbook wb;
		auto path = path_helper::get_data_directory("reader/charset-excel.xlsx");
		wb.load(path);
        
        auto ws = wb["Sheet1"];
        auto val = ws.get_cell("A1").get_value<std::string>();
        TS_ASSERT_EQUALS(val, "DirenÃ§");
    }
    
    void test_read_shared_strings_max_range()
    {
        xlnt::workbook wb;
		const auto path = path_helper::get_data_directory("reader/shared_strings-max_range.xlsx");
		wb.load(path);

        auto ws = wb["Sheet1"];
        auto val = ws.get_cell("A1").get_value<std::string>();
        TS_ASSERT_EQUALS(val, "Donald");
    }
    
    void test_read_shared_strings_multiple_r_nodes()
    {
        auto path = path_helper::get_data_directory("reader/shared_strings-multiple_r_nodes.xlsx");

        xlnt::workbook wb;
		wb.load(path);

        auto ws = wb["Sheet1"];
        auto val = ws.get_cell("A1").get_value<std::string>();
        TS_ASSERT_EQUALS(val, "abcdef");
    }

    xlnt::workbook date_mac_1904()
    {
        xlnt::workbook wb;
		wb.load(path_helper::get_data_directory("reader/date_1904.xlsx"));

        return wb;
    }
    
    xlnt::workbook date_std_1900()
    {
        xlnt::workbook wb;
		wb.load(path_helper::get_data_directory("reader/date_1900.xlsx"));

        return wb;
    }
    
    void test_read_win_base_date()
    {
        auto wb = date_std_1900();
        TS_ASSERT_EQUALS(wb.get_base_date(), xlnt::calendar::windows_1900);
    }
    
    void test_read_mac_base_date()
    {
        auto wb = date_mac_1904();
        TS_ASSERT_EQUALS(wb.get_base_date(), xlnt::calendar::mac_1904);
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
    
    void test_read_no_theme()
    {
        auto path = path_helper::get_data_directory("genuine/libreoffice_nrt.xlsx");
        
        xlnt::workbook wb;
        TS_ASSERT_THROWS_NOTHING(wb.load(path));
    }
    
    void _test_read_complex_formulae()
    {
        /*
        auto path = PathHelper::GetDataDirectory("reader/formulae.xlsx");
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
        auto path = path_helper::get_data_directory("reader/formulae.xlsx");
        
        xlnt::workbook wb;
		wb.set_data_only(true);
		wb.load(path);

        auto ws = wb.get_active_sheet();

        TS_ASSERT(ws.get_formula_attributes().empty());
        TS_ASSERT(ws.get_workbook().get_data_only());
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
        
        auto path = path_helper::get_data_directory("reader/contains_chartsheets.xlsx");

        xlnt::workbook wb;
        wb.load(path);
        
        auto &result = wb.get_manifest().get_parts();

        if(result.size() != expected.size())
        {
            TS_ASSERT_EQUALS(result.size(), expected.size());
            return;
        }

        for(std::size_t i = 0; i < expected.size(); i++)
        {
            TS_ASSERT(wb.get_manifest().has_part(xlnt::path(expected[i].first)));
            TS_ASSERT_EQUALS(wb.get_manifest().get_part_content_type_string(xlnt::path(expected[i].first)), expected[i].second);
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
            auto path = path_helper::get_data_directory("genuine/guess_types.xlsx");
            
            xlnt::workbook wb;
			wb.set_guess_types(guess);
			wb.load(path);

            auto ws = wb.get_active_sheet();
            TS_ASSERT(ws.get_cell("D2").get_data_type() == dtype);
        }
    }
    
    void test_read_autofilter()
    {
        auto path = path_helper::get_data_directory("reader/bug275.xlsx");

        xlnt::workbook wb;
		wb.load(path);
        
        auto ws = wb.get_active_sheet();
        TS_ASSERT_EQUALS(ws.get_auto_filter().to_string(), "A1:B6");
    }
    
    void test_bad_formats_xlsb()
    {
        auto path = path_helper::get_data_directory("genuine/a.xlsb");
        xlnt::workbook wb;
        TS_ASSERT_THROWS(wb.load(path), xlnt::invalid_file);
    }
    
    void test_bad_formats_xls()
    {
        auto path = path_helper::get_data_directory("genuine/a.xls");
        xlnt::workbook wb;
		TS_ASSERT_THROWS(wb.load(path), xlnt::invalid_file);
    }
    
    void test_bad_formats_no()
    {
        auto path = path_helper::get_data_directory("genuine/a.no-format");
        xlnt::workbook wb;
		TS_ASSERT_THROWS(wb.load(path), xlnt::invalid_file);
    }


    void test_read_shared_strings_with_runs()
    {
		xlnt::workbook wb;
		wb.load("reader/genuine/a.xlsx");

		const auto &strings = wb.get_shared_strings();

        TS_ASSERT_EQUALS(strings.size(), 1);
        TS_ASSERT_EQUALS(strings.front().get_runs().size(), 1);
        TS_ASSERT_EQUALS(strings.front().get_runs().front().get_size(), 13);
        TS_ASSERT_EQUALS(strings.front().get_runs().front().get_color(), "color");
        TS_ASSERT_EQUALS(strings.front().get_runs().front().get_font(), "font");
        TS_ASSERT_EQUALS(strings.front().get_runs().front().get_family(), 12);
        TS_ASSERT_EQUALS(strings.front().get_runs().front().get_scheme(), "scheme");
    }

    void test_read_inlinestr()
    {
        xlnt::workbook wb;
        wb.load(path_helper::get_data_directory("genuine/empty.xlsx"));
        TS_ASSERT_EQUALS(wb.get_sheet_by_index(0).get_cell("A1").get_value<std::string>(), "This is cell A1 in Sheet 1");
    }

    void test_determine_document_type()
    {
        xlnt::workbook wb;
        TS_ASSERT_THROWS(wb.load(path_helper::get_data_directory("1_empty.txt")), xlnt::invalid_file);
        TS_ASSERT_THROWS(wb.load(path_helper::get_data_directory("2_not-empty.txt")), xlnt::invalid_file);
        TS_ASSERT_THROWS(wb.load(path_helper::get_data_directory("3_empty.zip")), xlnt::invalid_file);
        TS_ASSERT_THROWS(wb.load(path_helper::get_data_directory("4_not-package.zip")), xlnt::invalid_file);
        TS_ASSERT_THROWS(wb.load(path_helper::get_data_directory("5_visio.vsdx")), xlnt::invalid_file);
        TS_ASSERT_THROWS(wb.load(path_helper::get_data_directory("6_document.docx")), xlnt::invalid_file);
        TS_ASSERT_THROWS(wb.load(path_helper::get_data_directory("7_presentation.pptx")), xlnt::invalid_file);
    }
};
