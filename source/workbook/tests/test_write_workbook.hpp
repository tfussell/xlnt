#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <detail/workbook_serializer.hpp>
#include <helpers/path_helper.hpp>
#include <xlnt/utils/exceptions.hpp>

class test_write_workbook : public CxxTest::TestSuite
{
public:
    void test_write_auto_filter()
    {
        xlnt::workbook wb;
        auto ws = wb.create_sheet();
        ws.get_cell("F42").set_value("hello");
        ws.auto_filter("A1:F1");
    
        xlnt::workbook_serializer serializer(wb);
        pugi::xml_document xml;
        serializer.write_workbook(xml);

        auto diff = Helper::compare_xml(PathHelper::read_file("workbook_auto_filter.xml"), xml);
        TS_ASSERT(!diff);
    }
    
    void test_write_hidden_worksheet()
    {
        xlnt::workbook wb;
        auto ws = wb.create_sheet();
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
        
        auto diff = Helper::compare_xml(expected, xml);
        TS_ASSERT(!diff);
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
        TemporaryFile file;

        xlnt::excel_serializer serializer(wb);
        serializer.save_workbook(file.GetFilename());
        
        TS_ASSERT(PathHelper::FileExists(file.GetFilename()));
    }
    
    void test_write_virtual_workbook()
    {
        xlnt::workbook old_wb, new_wb;
        
        xlnt::excel_serializer serializer(old_wb);
        std::vector<std::uint8_t> wb_bytes;
        serializer.save_virtual_workbook(wb_bytes);
        
        xlnt::excel_serializer deserializer(new_wb);
        deserializer.load_virtual_workbook(wb_bytes);

        TS_ASSERT(new_wb != nullptr);
    }
    
    void test_write_workbook_rels()
    {
        xlnt::workbook wb;
        xlnt::zip_file archive;
        xlnt::relationship_serializer serializer(archive);
        serializer.write_relationships(wb.get_relationships(), "xl/workbook.xml");
        pugi::xml_document observed;
        observed.load(archive.read("xl/_rels/workbook.xml.rels").c_str());
        auto filename = "workbook.xml.rels";
        auto diff = Helper::compare_xml(PathHelper::read_file(filename), observed);
        TS_ASSERT(!diff);
    }
    
    void test_write_workbook_()
    {
        xlnt::workbook wb;
        xlnt::workbook_serializer serializer(wb);
        pugi::xml_document xml;
        serializer.write_workbook(xml);
        auto filename = PathHelper::GetDataDirectory("/workbook.xml");
        auto diff = Helper::compare_xml(PathHelper::read_file(filename), xml);
        TS_ASSERT(!diff);
    }
    
    void test_write_named_range()
    {
        xlnt::workbook wb;
        auto ws = wb.create_sheet();
        wb.create_named_range("test_range", ws, "A1:B5");
        xlnt::workbook_serializer serializer(wb);
        pugi::xml_document xml;
        serializer.write_named_ranges(xml.root());
        std::string expected =
        "<root>"
            "<s:definedName xmlns:s=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" name=\"test_range\">'Sheet'!$A$1:$B$5</s:definedName>"
        "</root>";
        auto diff = Helper::compare_xml(expected, xml);
        TS_ASSERT(!diff);
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
        auto diff = Helper::compare_xml(expected, xml);
        TS_ASSERT(!diff);
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
        
        auto diff = Helper::compare_xml(expected, observed);
        TS_ASSERT(diff);
    }
              
private:
    xlnt::workbook wb_;
};
