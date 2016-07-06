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
    void test_write_number_formats()
    {
        xlnt::workbook wb;
        wb.get_active_sheet().get_cell("A1").set_number_format(xlnt::number_format("YYYY"));
        xlnt::excel_serializer excel_serializer(wb);
        xlnt::style_serializer style_serializer(excel_serializer.get_stylesheet());
        pugi::xml_document observed;
        style_serializer.write_stylesheet(observed);
        pugi::xml_document expected_doc;
        std::string expected =
        "    <numFmts count=\"1\">"
        "    <numFmt formatCode=\"YYYY\" numFmtId=\"164\"></numFmt>"
        "    </numFmts>";
        expected_doc.load(expected.c_str());
        auto diff = Helper::compare_xml(expected_doc.child("numFmts"), observed.child("styleSheet").child("numFmts"));
        TS_ASSERT(diff);
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

        xlnt::excel_serializer e(wb);
        xlnt::style_serializer serializer(e.get_stylesheet());
        pugi::xml_document xml;
        serializer.write_stylesheet(xml);

        TS_ASSERT(Helper::compare_xml(PathHelper::GetDataDirectory("/writer/expected/simple-styles.xml"), xml));
    }
    
    void test_empty_workbook()
    {
        xlnt::workbook wb;
        xlnt::excel_serializer e(wb);
        xlnt::style_serializer serializer(e.get_stylesheet());
        auto expected =
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
        pugi::xml_document xml;
        serializer.write_stylesheet(xml);
        TS_ASSERT(Helper::compare_xml(expected, xml));
    }
};
