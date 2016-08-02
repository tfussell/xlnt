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
		observed.save(std::cout);
        auto comparison = xml_helper::compare_xml(expected.root(), observed.root());
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

		xlnt::border b;
		xlnt::border::border_property prop;
		prop.set_style(xlnt::border_style::dashdot);
		prop.set_color(xlnt::rgb_color("ffff0000"));
		b.set_side(xlnt::border::side::top, prop);
		ws.get_cell("D10").set_border(b);

        std::string expected = path_helper::read_file(path_helper::get_data_directory("/writer/expected/simple-styles.xml"));
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

    void test_alignment()
    {
        xlnt::workbook wb;
        xlnt::alignment a;
        a.set_horizontal(xlnt::horizontal_alignment::center_continuous);
        a.set_vertical(xlnt::vertical_alignment::justify);
        a.set_wrap_text(true);
        a.set_shrink_to_fit(true);
        wb.get_active_sheet().get_cell("A1").set_alignment(a);

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
        "  <cellXfs count=\"2\">"
        "    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>"
        "    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\">"
        "      <alignment vertical=\"justify\" horizontal=\"center-continuous\" wrapText=\"1\" shrinkToFit=\"1\"/>"
        "    </xf>"
        "  </cellXfs>"
        "  <cellStyles count=\"1\">"
        "    <cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>"
        "  </cellStyles>"
        "  <dxfs count=\"0\"/>"
        "  <tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium9\" defaultPivotStyle=\"PivotStyleMedium7\"/>"
        "</styleSheet>";
        TS_ASSERT(style_xml_matches(expected, wb));
    }

    void test_fill()
    {
        xlnt::workbook wb;

        xlnt::fill pattern_fill = xlnt::fill::pattern(xlnt::pattern_fill::type::solid);
        pattern_fill.get_pattern_fill().set_foreground_color(xlnt::color::red());
        pattern_fill.get_pattern_fill().set_background_color(xlnt::color::blue());
        wb.get_active_sheet().get_cell("A1").set_fill(pattern_fill);

        xlnt::fill gradient_fill_linear = xlnt::fill::gradient(xlnt::gradient_fill::type::linear);
        gradient_fill_linear.get_gradient_fill().set_degree(90);
        wb.get_active_sheet().get_cell("A1").set_fill(gradient_fill_linear);

        xlnt::fill gradient_fill_path = xlnt::fill::gradient(xlnt::gradient_fill::type::path);
        gradient_fill_path.get_gradient_fill().set_gradient_left(1);
        gradient_fill_path.get_gradient_fill().set_gradient_right(2);
        gradient_fill_path.get_gradient_fill().set_gradient_top(3);
        gradient_fill_path.get_gradient_fill().set_gradient_bottom(4);
        gradient_fill_path.get_gradient_fill().add_stop(0, xlnt::color::red());
        gradient_fill_path.get_gradient_fill().add_stop(1, xlnt::color::blue());
        wb.get_active_sheet().get_cell("A1").set_fill(gradient_fill_path);

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
        "  <fills count=\"5\">"
        "    <fill>"
        "      <patternFill patternType=\"none\"/>"
        "    </fill>"
        "    <fill>"
        "      <patternFill patternType=\"gray125\"/>"
        "    </fill>"
        "    <fill>"
        "      <patternFill patternType=\"solid\">"
        "        <fgColor rgb=\"ffff0000\"/>"
        "        <bgColor rgb=\"ff0000ff\"/>"
        "      </patternFill>"
        "    </fill>"
        "    <fill>"
        "      <gradientFill gradientType=\"linear\" degree=\"90\"/>"
        "    </fill>"
        "    <fill>"
        "      <gradientFill gradientType=\"path\" left=\"1\" right=\"2\" top=\"3\" bottom=\"4\">"
        "        <stop position=\"0\">"
        "          <color rgb=\"ff00ff00\"/>"
        "        </stop>"
        "        <stop position=\"1\">"
        "          <color rgb=\"ffffff00\"/>"
        "        </stop>"
        "      </gradientFill>"
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
        "    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"2\" borderId=\"0\" xfId=\"0\"/>"
        "    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"3\" borderId=\"0\" xfId=\"0\"/>"
        "    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"4\" borderId=\"0\" xfId=\"0\"/>"
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
