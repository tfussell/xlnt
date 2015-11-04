#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/serialization/relationship_serializer.hpp>
#include <xlnt/serialization/shared_strings_serializer.hpp>
#include <xlnt/serialization/workbook_serializer.hpp>
#include <xlnt/serialization/worksheet_serializer.hpp>
#include <xlnt/workbook/workbook.hpp>

#include "helpers/temporary_file.hpp"
#include "helpers/path_helper.hpp"
#include "helpers/helper.hpp"

class test_write : public CxxTest::TestSuite
{
public:
    void test_write_empty_workbook()
    {
        xlnt::workbook wbk;
        wbk.get_active_sheet().get_cell("A2").set_value("xlnt");
        wbk.get_active_sheet().get_cell("B5").set_value(88);
        wbk.get_active_sheet().get_cell("B5").set_number_format(xlnt::number_format::percentage_00());
        wbk.save(temp_file.GetFilename());
        
        if(PathHelper::FileExists(temp_file.GetFilename()))
        {
            PathHelper::DeleteFile(temp_file.GetFilename());
        }

        TS_ASSERT(!PathHelper::FileExists(temp_file.GetFilename()));
        wb_.save(temp_file.GetFilename());
        TS_ASSERT(PathHelper::FileExists(temp_file.GetFilename()));
    }

    void test_write_virtual_workbook()
    {
        xlnt::workbook old_wb;
        std::vector<unsigned char> saved_wb;
        TS_ASSERT(old_wb.save(saved_wb));
        xlnt::workbook new_wb;
        TS_ASSERT(new_wb.load(saved_wb));
    }

    void test_write_workbook_rels()
    {
        xlnt::workbook wb;
        xlnt::zip_file archive;
        xlnt::relationship_serializer serializer(archive);
        serializer.write_relationships(wb.get_relationships(), "xl/workbook.xml");
        auto content = xlnt::xml_serializer::deserialize(archive.read("xl/_rels/workbook.xml.rels"));

        TS_ASSERT(Helper::compare_xml(PathHelper::GetDataDirectory() + "/writer/expected/workbook.xml.rels", content));
    }

    void test_write_workbook()
    {
        xlnt::workbook wb;
        xlnt::workbook_serializer serializer(wb);
        auto content = serializer.write_workbook();

        TS_ASSERT(Helper::compare_xml(PathHelper::GetDataDirectory() + "/writer/expected/workbook.xml", content));
    }

    void test_write_string_table()
    {
        xlnt::workbook wb;
        auto ws = wb.get_active_sheet();
        
        ws.get_cell("A1").set_value("hello");
        ws.get_cell("A2").set_value("world");
        ws.get_cell("A3").set_value("nice");
        
        auto content = xlnt::shared_strings_serializer::write_shared_strings(wb.get_shared_strings());
        
        TS_ASSERT(Helper::compare_xml(PathHelper::GetDataDirectory() + "/writer/expected/sharedStrings.xml", content));
    }

    void test_write_worksheet()
    {
		auto ws = wb_.create_sheet();
        ws.get_cell("F42").set_value("hello");
        
        xlnt::worksheet_serializer serializer(ws);
        auto observed = serializer.write_worksheet();
    
        TS_ASSERT(Helper::compare_xml(PathHelper::GetDataDirectory() + "/writer/expected/sheet1.xml", observed));
    }

    void test_write_hidden_worksheet()
    {
		auto ws = wb_.create_sheet();
        ws.get_page_setup().set_sheet_state(xlnt::page_setup::sheet_state::hidden);
        ws.get_cell("F42").set_value("hello");

        xlnt::worksheet_serializer serializer(ws);
        auto observed = serializer.write_worksheet();
        
        TS_ASSERT(Helper::compare_xml(PathHelper::GetDataDirectory() + "/writer/expected/sheet1.xml", observed));
    }

    void test_write_bool()
    {
        auto ws = wb_.create_sheet();
        ws.get_cell("F42").set_value(false);
        ws.get_cell("F43").set_value(true);

        xlnt::worksheet_serializer serializer(ws);
        auto observed = serializer.write_worksheet();
        
        TS_ASSERT(Helper::compare_xml(PathHelper::GetDataDirectory() + "/writer/expected/sheet1_bool.xml", observed));
    }

    void test_write_formula()
    {
		auto ws = wb_.create_sheet();
        ws.get_cell("F1").set_value(10);
        ws.get_cell("F2").set_value(32);
        ws.get_cell("F3").set_formula("F1+F2");

        xlnt::worksheet_serializer serializer(ws);
        auto observed = serializer.write_worksheet();
        
        TS_ASSERT(Helper::compare_xml(PathHelper::GetDataDirectory() + "/writer/expected/sheet1_formula.xml", observed));
    }

    void test_write_height()
    {
		auto ws = wb_.create_sheet();
        ws.get_cell("F1").set_value(10);
        ws.get_row_properties(ws.get_cell("F1").get_row()).height = 30;

        xlnt::worksheet_serializer serializer(ws);
        auto observed = serializer.write_worksheet();
        
        TS_ASSERT(Helper::compare_xml(PathHelper::GetDataDirectory() + "/writer/expected/sheet1_height.xml", observed));
    }

    void test_write_hyperlink()
    {
        xlnt::workbook clean_wb;
        auto ws = clean_wb.get_active_sheet();
        
        ws.get_cell("A1").set_value("test");
        ws.get_cell("A1").set_hyperlink("http://test.com");

        xlnt::worksheet_serializer serializer(ws);
        auto observed = serializer.write_worksheet();
        
        TS_ASSERT(Helper::compare_xml(PathHelper::GetDataDirectory() + "/writer/expected/sheet1_hyperlink.xml", observed));
    }

    void test_write_hyperlink_rels()
    {
		auto ws = wb_.create_sheet();
        TS_ASSERT_EQUALS(0, ws.get_relationships().size());
        ws.get_cell("A1").set_value("test");
        ws.get_cell("A1").set_hyperlink("http://test.com/");
        TS_ASSERT_EQUALS(1, ws.get_relationships().size());
        ws.get_cell("A2").set_value("test");
        ws.get_cell("A2").set_hyperlink("http://test2.com/");
        TS_ASSERT_EQUALS(2, ws.get_relationships().size());
        
        xlnt::zip_file archive;
        xlnt::relationship_serializer serializer(archive);
        serializer.write_relationships(ws.get_relationships(), "xl/worksheets/sheet1.xml");
        auto content = xlnt::xml_serializer::deserialize(archive.read("xl/worksheets/_rels/sheet1.xml.rels"));
        
        TS_ASSERT(Helper::compare_xml(PathHelper::GetDataDirectory() + "/writer/expected/sheet1_hyperlink.xml.rels", content));
    }
    
    void _test_write_hyperlink_image_rels()
    {
        TS_SKIP("not implemented");
    }

    void test_hyperlink_value()
    {
		auto ws = wb_.create_sheet();
        ws.get_cell("A1").set_hyperlink("http://test.com");
        TS_ASSERT_EQUALS("http://test.com", ws.get_cell("A1").get_value<xlnt::string>());
        ws.get_cell("A1").set_value("test");
        TS_ASSERT_EQUALS("test", ws.get_cell("A1").get_value<xlnt::string>());
    }

    void test_write_auto_filter()
    {
        xlnt::workbook wb;
        auto ws = wb.get_sheet_by_index(0);
        ws.get_cell("F42").set_value("hello");
        ws.auto_filter("A1:F1");

        xlnt::worksheet_serializer ws_serializer(ws);
        auto observed = ws_serializer.write_worksheet();
        
        TS_ASSERT(Helper::compare_xml(PathHelper::GetDataDirectory() + "/writer/expected/sheet1_auto_filter.xml", observed));

        xlnt::workbook_serializer wb_serializer(wb);
        auto observed2 = wb_serializer.write_workbook();

        TS_ASSERT(Helper::compare_xml(PathHelper::GetDataDirectory() + "/writer/expected/workbook_auto_filter.xml", observed2));
    }
    
    void test_write_auto_filter_filter_column()
    {
        
    }
    
    void test_write_auto_filter_sort_condition()
    {
        
    }

    void test_freeze_panes_horiz()
    {
        auto ws = wb_.create_sheet();
        ws.get_cell("F42").set_value("hello");
        ws.freeze_panes("A4");

        xlnt::worksheet_serializer serializer(ws);
        auto observed = serializer.write_worksheet();
        
        TS_ASSERT(Helper::compare_xml(PathHelper::GetDataDirectory() + "/writer/expected/sheet1_freeze_panes_horiz.xml", observed));
    }

    void test_freeze_panes_vert()
    {
		auto ws = wb_.create_sheet();
        ws.get_cell("F42").set_value("hello");
        ws.freeze_panes("D1");

        xlnt::worksheet_serializer serializer(ws);
        auto observed = serializer.write_worksheet();
        
        TS_ASSERT(Helper::compare_xml(PathHelper::GetDataDirectory() + "/writer/expected/sheet1_freeze_panes_vert.xml", observed));
    }

    void test_freeze_panes_both()
    {
		auto ws = wb_.create_sheet();
        ws.get_cell("F42").set_value("hello");
        ws.freeze_panes("D4");
        
        xlnt::worksheet_serializer serializer(ws);
        auto observed = serializer.write_worksheet();
        
        TS_ASSERT(Helper::compare_xml(PathHelper::GetDataDirectory() + "/writer/expected/sheet1_freeze_panes_both.xml", observed));
    }

    void test_long_number()
    {
		auto ws = wb_.create_sheet();
        ws.get_cell("A1").set_value(9781231231230LL);
   
        xlnt::worksheet_serializer serializer(ws);
        auto observed = serializer.write_worksheet();
        
        TS_ASSERT(Helper::compare_xml(PathHelper::GetDataDirectory() + "/writer/expected/long_number.xml", observed));
    }

    void test_short_number()
    {
		auto ws = wb_.create_sheet();
        ws.get_cell("A1").set_value(1234567890);
        
        xlnt::worksheet_serializer serializer(ws);
        auto observed = serializer.write_worksheet();
        
        TS_ASSERT(Helper::compare_xml(PathHelper::GetDataDirectory() + "/writer/expected/short_number.xml", observed));
    }
    
    void _test_write_images()
    {
        TS_SKIP("not implemented");
    }

 private:
    TemporaryFile temp_file;
    xlnt::workbook wb_;
};
