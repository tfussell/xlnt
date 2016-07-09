#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <detail/style_serializer.hpp>
#include <detail/stylesheet.hpp>
#include <detail/workbook_impl.hpp>
#include <helpers/path_helper.hpp>

class test_style_writer : public CxxTest::TestSuite
{
public:
    bool style_xml_matches(const std::string &expected_string, xlnt::workbook &wb)
    {
        xlnt::excel_serializer excel_serializer(wb);
        xlnt::style_serializer style_serializer(excel_serializer.get_stylesheet());
        pugi::xml_document observed;
        style_serializer.write_stylesheet(observed);
        pugi::xml_document expected;
        expected.load(expected_string.c_str());

        auto comparison = Helper::compare_xml(expected.root(), observed.root());
        return (bool)comparison;
    }

    void test_write_custom_number_format()
    {
        xlnt::workbook wb;
        wb.get_active_sheet().get_cell("A1").set_number_format(xlnt::number_format("YYYY"));
        auto expected =
        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" mc:Ignorable=\"x14ac\">"
        "  <numFmts count=\"1\">"
        "    <numFmt formatCode=\"YYYY\" numFmtId=\"164\"></numFmt>"
        "  </numFmts>"
        "  <fonts count=\"1\">"
        "    <font>"
        "      <sz val=\"12\"/>"
        "      <color theme=\"1\"/>"
        "      <name val=\"Calibri\"/>"
        "      <family val=\"2\"/>"
        "      <scheme val=\"minor\"/>"
        "    </font>"
        "  </fonts>"
        "  <fills count=\"2\">"
        "    <fill>"
        "      <patternFill patternType=\"none\"/>"
        "    </fill>"
        "    <fill>"
        "      <patternFill patternType=\"gray125\"/>"
        "    </fill>"
        "  </fills>"
        "  <borders count=\"1\">"
        "    <border>"
        "      <left/>"
        "      <right/>"
        "      <top/>"
        "      <bottom/>"
        "      <diagonal/>"
        "    </border>"
        "  </borders>"
        "  <cellStyleXfs count=\"1\">"
        "    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>"
        "  </cellStyleXfs>"
        "  <cellXfs count=\"2\">"
        "    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>"
        "    <xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>"
        "  </cellXfs>"
        "  <cellStyles count=\"1\">"
        "    <cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>"
        "  </cellStyles>"
        "  <dxfs count=\"0\"/>"
        "  <tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium9\" defaultPivotStyle=\"PivotStyleMedium7\"/>"
        "</styleSheet>";
        TS_ASSERT(style_xml_matches(expected, wb));
    }

    void test_simple_styles()
    {
        xlnt::workbook wb;
        wb.set_guess_types(true);
        auto ws = wb.get_active_sheet();
        ws.get_cell("A1").set_value("12.34%");
        auto now = xlnt::date::today();
        ws.get_cell("A2").set_value(now);
        ws.get_cell("A3").set_value("This is a test");
        ws.get_cell("A4").set_value("31.31415");
        ws.get_cell("A5");
        
        ws.get_cell("D9").set_number_format(xlnt::number_format::number_00());
        xlnt::protection locked(true, false);
        ws.get_cell("D9").set_protection(locked);
        xlnt::protection hidden(true, true);
        ws.get_cell("E1").set_protection(hidden);

        std::string expected = PathHelper::read_file(PathHelper::GetDataDirectory("/writer/expected/simple-styles.xml"));
        TS_ASSERT(style_xml_matches(expected, wb));
    }
    
    void test_empty_workbook()
    {
        xlnt::workbook wb;
        std::string expected =
        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" mc:Ignorable=\"x14ac\">"
        "  <fonts count=\"1\">"
        "    <font>"
        "      <sz val=\"12\"/>"
        "      <color theme=\"1\"/>"
        "      <name val=\"Calibri\"/>"
        "      <family val=\"2\"/>"
        "      <scheme val=\"minor\"/>"
        "    </font>"
        "  </fonts>"
        "  <fills count=\"2\">"
        "    <fill>"
        "      <patternFill patternType=\"none\"/>"
        "    </fill>"
        "    <fill>"
        "      <patternFill patternType=\"gray125\"/>"
        "    </fill>"
        "  </fills>"
        "  <borders count=\"1\">"
        "    <border>"
        "      <left/>"
        "      <right/>"
        "      <top/>"
        "      <bottom/>"
        "      <diagonal/>"
        "    </border>"
        "  </borders>"
        "  <cellStyleXfs count=\"1\">"
        "    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>"
        "  </cellStyleXfs>"
        "  <cellXfs count=\"1\">"
        "    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>"
        "  </cellXfs>"
        "  <cellStyles count=\"1\">"
        "    <cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>"
        "  </cellStyles>"
        "  <dxfs count=\"0\"/>"
        "  <tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium9\" defaultPivotStyle=\"PivotStyleMedium7\"/>"
        "</styleSheet>";
        TS_ASSERT(style_xml_matches(expected, wb));
    }
    
    void test_complex_font()
    {
        xlnt::font f;
        f.set_bold(true);
        f.set_color(xlnt::color::red());
        f.set_family(3);
        f.set_italic(true);
        f.set_name("Consolas");
        f.set_scheme("major");
        f.set_size(21);
        f.set_strikethrough(true);
        f.set_underline(xlnt::font::underline_style::double_accounting);

        xlnt::workbook wb;
        wb.get_active_sheet().get_cell("A1").set_font(f);

        std::string expected =
        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" mc:Ignorable=\"x14ac\">"
        "  <fonts count=\"2\">"
        "    <font>"
        "      <sz val=\"12\"/>"
        "      <color theme=\"1\"/>"
        "      <name val=\"Calibri\"/>"
        "      <family val=\"2\"/>"
        "      <scheme val=\"minor\"/>"
        "    </font>"
        "    <font>"
        "      <b val=\"1\"/>"
        "      <i val=\"1\"/>"
        "      <u val=\"double-accounting\"/>"
        "      <strike val=\"1\"/>"
        "      <sz val=\"21\"/>"
        "      <color rgb=\"ffff0000\"/>"
        "      <name val=\"Consolas\"/>"
        "      <family val=\"3\"/>"
        "      <scheme val=\"major\"/>"
        "    </font>"
        "  </fonts>"
        "  <fills count=\"2\">"
        "    <fill>"
        "      <patternFill patternType=\"none\"/>"
        "    </fill>"
        "    <fill>"
        "      <patternFill patternType=\"gray125\"/>"
        "    </fill>"
        "  </fills>"
        "  <borders count=\"1\">"
        "    <border>"
        "      <left/>"
        "      <right/>"
        "      <top/>"
        "      <bottom/>"
        "      <diagonal/>"
        "    </border>"
        "  </borders>"
        "  <cellStyleXfs count=\"1\">"
        "    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>"
        "  </cellStyleXfs>"
        "  <cellXfs count=\"2\">"
        "    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>"
        "    <xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>"
        "  </cellXfs>"
        "  <cellStyles count=\"1\">"
        "    <cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>"
        "  </cellStyles>"
        "  <dxfs count=\"0\"/>"
        "  <tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium9\" defaultPivotStyle=\"PivotStyleMedium7\"/>"
        "</styleSheet>";

        TS_ASSERT(style_xml_matches(expected, wb));
    }
};
