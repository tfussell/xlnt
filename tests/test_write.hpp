#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/xlnt.hpp>
#include "helpers/temporary_file.hpp"
#include "helpers/path_helper.hpp"
#include "helpers/helper.hpp"

class test_write : public CxxTest::TestSuite
{
public:
    void test_write_empty_workbook()
    {
        if(PathHelper::FileExists(temp_file.GetFilename()))
        {
            PathHelper::DeleteFile(temp_file.GetFilename());
        }

        TS_ASSERT(!PathHelper::FileExists(temp_file.GetFilename()));
        wb.save(temp_file.GetFilename());
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
        auto content = xlnt::writer::write_workbook_rels(wb);

        TS_ASSERT(Helper::EqualsFileContent(PathHelper::GetDataDirectory() + "/writer/expected/workbook.xml.rels", content));
    }

    void test_write_workbook()
    {
        xlnt::workbook wb;
        auto content = xlnt::writer::write_workbook(wb);

        TS_ASSERT(Helper::EqualsFileContent(PathHelper::GetDataDirectory() + "/writer/expected/workbook.xml", content));
    }

    void test_write_string_table()
    {
        std::vector<std::string> table = {"hello", "world", "nice"};
        auto content = xlnt::writer::write_string_table(table);
        TS_ASSERT(Helper::EqualsFileContent(PathHelper::GetDataDirectory() + "/writer/expected/sharedStrings.xml", content));
    }

    void test_write_worksheet()
    {
        auto ws = wb.create_sheet();
        ws.get_cell("F42") = "hello";
        auto content = xlnt::writer::write_worksheet(ws, {"hello"}, {});
        TS_ASSERT(Helper::EqualsFileContent(PathHelper::GetDataDirectory() + "/writer/expected/sheet1.xml", content));
    }

    void test_write_hidden_worksheet()
    {
        auto ws = wb.create_sheet();
        ws.get_page_setup().set_sheet_state(xlnt::page_setup::sheet_state::hidden);
        ws.get_cell("F42") = "hello";
        auto content = xlnt::writer::write_worksheet(ws, {"hello"}, {});
        TS_ASSERT(Helper::EqualsFileContent(PathHelper::GetDataDirectory() + "/writer/expected/sheet1.xml", content));
    }

    void test_write_bool()
    {
        auto ws = wb.create_sheet();
        ws.get_cell("F42") = false;
        ws.get_cell("F43") = true;
        auto content = xlnt::writer::write_worksheet(ws, {}, {});
	TS_ASSERT(Helper::EqualsFileContent(PathHelper::GetDataDirectory() + "/writer/expected/sheet1_bool.xml", content));
    }

    void test_write_formula()
    {
        auto ws = wb.create_sheet();
        ws.get_cell("F1") = 10;
        ws.get_cell("F2") = 32;
        ws.get_cell("F3") = "=F1+F2";
        auto content = xlnt::writer::write_worksheet(ws, {}, {});
        TS_ASSERT(Helper::EqualsFileContent(PathHelper::GetDataDirectory() + "/writer/expected/sheet1_formula.xml", content));
    }

    void test_write_style()
    {
        auto ws = wb.create_sheet();
        ws.get_cell("F1") = "13%";
        auto style_id_by_hash = xlnt::style_writer(wb).get_style_by_hash();
        auto content = xlnt::writer::write_worksheet(ws, {}, style_id_by_hash);
        TS_ASSERT(Helper::EqualsFileContent(PathHelper::GetDataDirectory() + "/writer/expected/sheet1_style.xml", content));
    }

    void test_write_height()
    {
        auto ws = wb.create_sheet();
        ws.get_cell("F1") = 10;
        ws.get_row_properties(ws.get_cell("F1").get_row()).set_height(30);
        auto content = xlnt::writer::write_worksheet(ws, {}, {});
        TS_ASSERT(Helper::EqualsFileContent(PathHelper::GetDataDirectory() + "/writer/expected/sheet1_height.xml", content));
    }

    void test_write_hyperlink()
    {
        auto ws = wb.create_sheet();
        ws.get_cell("A1") = "test";
        ws.get_cell("A1").set_hyperlink("http://test.com");
        auto content = xlnt::writer::write_worksheet(ws, {"test"}, {});
        TS_ASSERT(Helper::EqualsFileContent(PathHelper::GetDataDirectory() + "/writer/expected/sheet1_hyperlink.xml", content));
    }

    void test_write_hyperlink_rels()
    {
        auto ws = wb.create_sheet();
        TS_ASSERT_EQUALS(0, ws.get_relationships().size());
        ws.get_cell("A1") = "test";
        ws.get_cell("A1").set_hyperlink("http://test.com/");
        TS_ASSERT_EQUALS(1, ws.get_relationships().size());
        ws.get_cell("A2") = "test";
        ws.get_cell("A2").set_hyperlink("http://test2.com/");
        TS_ASSERT_EQUALS(2, ws.get_relationships().size());
        auto content = xlnt::writer::write_worksheet_rels(ws, 1, 1);
        TS_ASSERT(Helper::EqualsFileContent(PathHelper::GetDataDirectory() + "/writer/expected/sheet1_hyperlink.xml.rels", content));
    }

    void test_hyperlink_value()
    {
        auto ws = wb.create_sheet();
        ws.get_cell("A1").set_hyperlink("http://test.com");
        TS_ASSERT_EQUALS("http://test.com", ws.get_cell("A1"));
        ws.get_cell("A1") = "test";
        TS_ASSERT_EQUALS("test", ws.get_cell("A1"));
    }

    void test_write_auto_filter()
    {
        xlnt::workbook wb;
        auto ws = wb.get_sheet_by_index(0);
        ws.get_cell("F42") = "hello";
        ws.auto_filter("A1:F1");
        auto content = xlnt::writer::write_worksheet(ws, {"hello"}, {});
        TS_ASSERT(Helper::EqualsFileContent(PathHelper::GetDataDirectory() + "/writer/expected/sheet1_auto_filter.xml", content));

        content = xlnt::writer::write_workbook(wb);
        TS_ASSERT(Helper::EqualsFileContent(PathHelper::GetDataDirectory() + "/writer/expected/workbook_auto_filter.xml", content));
    }

    void test_freeze_panes_horiz()
    {
        auto ws = wb.create_sheet();
        ws.get_cell("F42") = "hello";
        ws.freeze_panes("A4");
        auto content = xlnt::writer::write_worksheet(ws, {"hello"}, {});
	TS_ASSERT(Helper::EqualsFileContent(PathHelper::GetDataDirectory() + "/writer/expected/sheet1_freeze_panes_horiz.xml", content));
    }

    void test_freeze_panes_vert()
    {
        auto ws = wb.create_sheet();
        ws.get_cell("F42") = "hello";
        ws.freeze_panes("D1");
        auto content = xlnt::writer::write_worksheet(ws, {"hello"}, {});
        TS_ASSERT(Helper::EqualsFileContent(PathHelper::GetDataDirectory() + "/writer/expected/sheet1_freeze_panes_vert.xml", content));
    }

    void test_freeze_panes_both()
    {
        auto ws = wb.create_sheet();
        ws.get_cell("F42") = "hello";
        ws.freeze_panes("D4");
        auto content = xlnt::writer::write_worksheet(ws, {"hello"}, {});
        TS_ASSERT(Helper::EqualsFileContent(PathHelper::GetDataDirectory() + "/writer/expected/sheet1_freeze_panes_both.xml", content));
    }

    void test_long_number()
    {
        auto ws = wb.create_sheet();
        ws.get_cell("A1") = 9781231231230;
        auto content = xlnt::writer::write_worksheet(ws, {}, {});
        TS_ASSERT(Helper::EqualsFileContent(PathHelper::GetDataDirectory() + "/writer/expected/long_number.xml", content));
    }

    void test_short_number()
    {
        auto ws = wb.create_sheet();
        ws.get_cell("A1") = 1234567890;
        auto content = xlnt::writer::write_worksheet(ws, {}, {});
        TS_ASSERT(Helper::EqualsFileContent(PathHelper::GetDataDirectory() + "/writer/expected/short_number.xml", content));
    }

 private:
    TemporaryFile temp_file;
    xlnt::workbook wb;
};
