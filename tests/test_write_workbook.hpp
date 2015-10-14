#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/common/exceptions.hpp>
#include <xlnt/writer/workbook_writer.hpp>

#include "helpers/path_helper.hpp"

namespace xlnt {

xlnt::workbook load_workbook(const std::vector<std::uint8_t> &bytes)
{
    return xlnt::workbook();
}

std::string write_workbook_rels(xlnt::workbook &wb)
{
    return "";
}

std::string write_root_rels(xlnt::workbook &wb)
{
    return "";
}

std::string write_defined_names(xlnt::workbook &wb)
{
    return "";
}
    
}

class test_write_workbook : public CxxTest::TestSuite
{
public:
    void test_write_auto_filter()
    {
        xlnt::workbook wb;
        auto ws = wb.create_sheet();
        ws.get_cell("F42").set_value("hello");
        ws.get_auto_filter() = "A1:F1";
    
        auto content = xlnt::write_workbook(wb);
        auto diff = Helper::compare_xml(PathHelper::read_file("workbook_auto_filter.xml"), content);
        TS_ASSERT(!diff);
    }
    
    void test_write_hidden_worksheet()
    {
        xlnt::workbook wb;
        auto ws = wb.create_sheet();
        ws.set_sheet_state(xlnt::page_setup::sheet_state::hidden);
        wb.create_sheet();
        auto xml = xlnt::write_workbook(wb);
        std::string expected =
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
        auto diff = Helper::compare_xml(xml, expected);
        TS_ASSERT(!diff);
    }
    
    void test_write_hidden_single_worksheet()
    {
        xlnt::workbook wb;
        auto ws = wb.get_active_sheet();
        ws.set_sheet_state(xlnt::page_setup::sheet_state::hidden);
        TS_ASSERT_THROWS(xlnt::write_workbook(wb), xlnt::value_error);
    }
    
    void test_write_empty_workbook()
    {
        xlnt::workbook wb;
        TemporaryFile file;
        xlnt::save_workbook(wb, file.GetFilename());
        TS_ASSERT(PathHelper::FileExists(file.GetFilename()));
    }
    
    void test_write_virtual_workbook()
    {
        xlnt::workbook old_wb;
        auto saved_wb = xlnt::save_virtual_workbook(old_wb);
        auto new_wb = xlnt::load_workbook(saved_wb);
        TS_ASSERT(new_wb != nullptr);
    }
    
    void test_write_workbook_rels()
    {
        xlnt::workbook wb;
        auto content = xlnt::write_workbook_rels(wb);
        auto filename = "workbook.xml.rels";
        auto diff = Helper::compare_xml(PathHelper::read_file(filename), content);
        TS_ASSERT(!diff);
    }
    
    void test_write_workbook_()
    {
        xlnt::workbook wb;
        auto content = xlnt::write_workbook(wb);
        auto filename = "workbook.xml";
        auto diff = Helper::compare_xml(PathHelper::read_file(filename), content);
        TS_ASSERT(!diff);
    }
    
    void test_write_named_range()
    {
        xlnt::workbook wb;
        auto ws = wb.create_sheet();
        xlnt::named_range xlrange("test_range", {{ws, "A1:B5"}});
        wb.add_named_range(xlrange);
        auto xml = xlnt::write_defined_names(wb);
        std::string expected =
        "<root>"
            "<s:definedName xmlns:s=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" name=\"test_range\">'Sheet'!$A$1:$B$5</s:definedName>"
        "</root>";
        auto diff = Helper::compare_xml(xml, expected);
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
    
        auto content = xlnt::write_workbook(wb);
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
        auto diff = Helper::compare_xml(content, expected);
        TS_ASSERT(!diff);
    }
    
    void test_write_root_rels()
    {
        xlnt::workbook wb;
        auto xml = xlnt::write_root_rels(wb);
        std::string expected =
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "    <Relationship Id=\"rId1\" Target=\"xl/workbook.xml\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\"/>"
        "    <Relationship Id=\"rId2\" Target=\"docProps/core.xml\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\"/>"
        "    <Relationship Id=\"rId3\" Target=\"docProps/app.xml\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\"/>"
        "</Relationships>";
        auto diff = Helper::compare_xml(xml, expected);
        TS_ASSERT(!diff);
    }
              
private:
    xlnt::workbook wb_;
};
