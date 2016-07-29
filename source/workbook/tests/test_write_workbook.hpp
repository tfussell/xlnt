#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <detail/shared_strings_serializer.hpp>
#include <detail/workbook_serializer.hpp>
#include <helpers/path_helper.hpp>
#include <xlnt/utils/exceptions.hpp>

class test_write_workbook : public CxxTest::TestSuite
{
public:
    void test_write_auto_filter()
    {
        xlnt::workbook wb;
        auto ws = wb.get_active_sheet();
        ws.get_cell("F42").set_value("hello");
        ws.auto_filter("A1:F1");
    
        xlnt::workbook_serializer serializer(wb);
        pugi::xml_document xml;
        serializer.write_workbook(xml);

        TS_ASSERT(xml_helper::compare_xml(path_helper::get_data_directory("/writer/expected/workbook_auto_filter.xml"), xml));
    }
    
    void test_write_hidden_worksheet()
    {
        xlnt::workbook wb;
        auto ws = wb.get_active_sheet();
        ws.set_sheet_state(xlnt::sheet_state::hidden);
        wb.create_sheet();

        xlnt::workbook_serializer serializer(wb);
        pugi::xml_document xml;
        serializer.write_workbook(xml);

        std::string expected_string =
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "    <workbookPr/>"
        "    <bookViews>"
        "        <workbookView activeTab=\"0\"/>"
        "    </bookViews>"
        "    <sheets>"
        "        <sheet name=\"Sheet\" sheetId=\"1\" state=\"hidden\" r:id=\"rId1\"/>"
        "        <sheet name=\"Sheet1\" sheetId=\"2\" r:id=\"rId2\"/>"
        "    </sheets>"
        "    <definedNames/>"
        "    <calcPr calcId=\"124519\" fullCalcOnLoad=\"1\"/>"
        "</workbook>";

        pugi::xml_document expected;
        expected.load(expected_string.c_str());

        TS_ASSERT(xml_helper::compare_xml(expected.child("workbook").child("sheets"),
            xml.child("workbook").child("sheets")));
    }
    
    void test_write_hidden_single_worksheet()
    {
        xlnt::workbook wb;
        auto ws = wb.get_active_sheet();
        ws.set_sheet_state(xlnt::sheet_state::hidden);
        
        xlnt::workbook_serializer serializer(wb);
        
        pugi::xml_document xml;
        TS_ASSERT_THROWS(serializer.write_workbook(xml), xlnt::value_error);
    }
    
    void test_write_empty_workbook()
    {
        xlnt::workbook wb;
        temporary_file file;

        xlnt::excel_serializer serializer(wb);
        serializer.save_workbook(file.get_filename());
        
        TS_ASSERT(path_helper::file_exists(file.get_filename()));
    }
    
    void test_write_virtual_workbook()
    {
        xlnt::workbook old_wb, new_wb;
        
        xlnt::excel_serializer serializer(old_wb);
        std::vector<std::uint8_t> wb_bytes;
        serializer.save_virtual_workbook(wb_bytes);
        
        xlnt::excel_serializer deserializer(new_wb);
        deserializer.load_virtual_workbook(wb_bytes);
    }
    
    void test_write_workbook_rels()
    {
        xlnt::workbook wb;
        auto ws = wb.get_active_sheet();
        ws.get_cell("A1").set_value("string");

        xlnt::zip_file archive;
        xlnt::relationship_serializer serializer(archive);
        serializer.write_relationships(wb.get_relationships(), "xl/workbook.xml");
        pugi::xml_document observed;
        observed.load(archive.read("xl/_rels/workbook.xml.rels").c_str());

        TS_ASSERT(xml_helper::compare_xml(path_helper::get_data_directory("/writer/expected/workbook.xml.rels"), observed));
    }
    
    void test_write_workbook_part()
    {
        xlnt::workbook wb;
        xlnt::workbook_serializer serializer(wb);
        pugi::xml_document xml;
        serializer.write_workbook(xml);
        auto filename = path_helper::get_data_directory("/writer/expected/workbook.xml");

        TS_ASSERT(xml_helper::compare_xml(filename, xml));
    }
    
    void test_write_named_range()
    {
        xlnt::workbook wb;
        auto ws = wb.get_active_sheet();
        wb.create_named_range("test_range", ws, "$A$1:$B$5");
        xlnt::workbook_serializer serializer(wb);
        pugi::xml_document xml;
        auto root = xml.root().append_child("root");
        serializer.write_named_ranges(root);
        std::string expected =
        "<root>"
            "<s:definedName xmlns:s=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" name=\"test_range\">'Sheet'!$A$1:$B$5</s:definedName>"
        "</root>";

        TS_ASSERT(xml_helper::compare_xml(expected, xml));
    }
    
    void test_read_workbook_code_name()
    {
//        with open(tmpl, "rb") as expected:
//        TS_ASSERT(read_workbook_code_name(expected.read()) == code_name
    }
    
    void test_write_workbook_code_name()
    {
        xlnt::workbook wb;
        wb.set_code_name("MyWB");
    
        xlnt::workbook_serializer serializer(wb);
        pugi::xml_document xml;
        serializer.write_workbook(xml);

        std::string expected =
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "    <workbookPr codeName=\"MyWB\"/>"
        "    <bookViews>"
        "        <workbookView activeTab=\"0\"/>"
        "    </bookViews>"
        "    <sheets>"
        "        <sheet name=\"Sheet\" sheetId=\"1\" r:id=\"rId1\"/>"
        "    </sheets>"
        "    <definedNames/>"
        "    <calcPr calcId=\"124519\" fullCalcOnLoad=\"1\"/>"
        "</workbook>";
        pugi::xml_document expected_xml;
        expected_xml.load(expected.c_str());

        TS_ASSERT(xml_helper::compare_xml(expected_xml.child("workbook").child("workbookPr"),
            xml.child("workbook").child("workbookPr")));
    }
    
    void test_write_root_rels()
    {
        xlnt::workbook wb;
        xlnt::zip_file archive;
        xlnt::relationship_serializer serializer(archive);
        serializer.write_relationships(wb.get_root_relationships(), "");
        pugi::xml_document observed;
        observed.load(archive.read("_rels/.rels").c_str());
        
        std::string expected =
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "    <Relationship Id=\"rId1\" Target=\"xl/workbook.xml\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\"/>"
        "    <Relationship Id=\"rId2\" Target=\"docProps/core.xml\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\"/>"
        "    <Relationship Id=\"rId3\" Target=\"docProps/app.xml\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\"/>"
        "</Relationships>";
        
        TS_ASSERT(xml_helper::compare_xml(expected, observed));
    }

    void test_write_shared_strings_with_runs()
    {
        xlnt::workbook wb;
        auto ws = wb.get_active_sheet();
        auto cell = ws.get_cell("A1");
        xlnt::text_run run;
        run.set_color("color");
        run.set_family(12);
        run.set_font("font");
        run.set_scheme("scheme");
        run.set_size(13);
        run.set_string("string");
        xlnt::text text;
        text.add_run(run);
        cell.set_value(text);

        pugi::xml_document xml;
        xlnt::shared_strings_serializer::write_shared_strings(wb.get_shared_strings(), xml);

        std::string expected =
        "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\" uniqueCount=\"1\">"
        "  <si>"
        "    <r>"
        "      <rPr>"
        "        <sz val=\"13\"/>"
        "        <color rgb=\"color\"/>"
        "        <rFont val=\"font\"/>"
        "        <family val=\"12\"/>"
        "        <scheme val=\"scheme\"/>"
        "      </rPr>"
        "      <t>string</t>"
        "    </r>"
        "  </si>"
        "</sst>";

        TS_ASSERT(xml_helper::compare_xml(expected, xml));
    }

    void test_write_worksheet_order()
    {
        auto path = path_helper::get_data_directory("/genuine/tab_order.xlsx");

        // Load an original workbook produced by Excel
        xlnt::workbook wb_src;
        {
          xlnt::excel_serializer serializer(wb_src);
          serializer.load_workbook(path);
        }

        // Save it to a new file, unmodified
        temporary_file file;
        {
          xlnt::excel_serializer serializer(wb_src);
          serializer.save_workbook(file.get_filename());
          TS_ASSERT(path_helper::file_exists(file.get_filename()));
        }

        // Load it again
        xlnt::workbook wb_dst;
        {
          xlnt::excel_serializer serializer(wb_dst);
          serializer.load_workbook(file.get_filename());
        }

        // Make sure the number of worksheets is the same
        auto count_src = std::distance(wb_src.begin(), wb_src.end());
        auto count_dst = std::distance(wb_dst.begin(), wb_dst.end());
        TS_ASSERT(count_src == count_dst);

        // Make sure the title of the first sheet matches
        auto ws1title_src = wb_src[0].get_title();
        auto ws1title_dst = wb_dst[0].get_title();
        TS_ASSERT(ws1title_src.compare(ws1title_dst) == 0);
    }

private:
    xlnt::workbook wb_;
};
