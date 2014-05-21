#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/xlnt.h>
#include "TemporaryFile.h"
#include "PathHelper.h"

class WriteTestSuite : public CxxTest::TestSuite
{
public:
    WriteTestSuite()
    {

    }

    void test_write_empty_workbook()
    {
        TemporaryFile temp_file;
        xlnt::workbook wb;
        auto dest_filename = temp_file.GetFilename();
        wb.save(dest_filename);
        TS_ASSERT(PathHelper::FileExists(dest_filename));
    }

    void test_write_virtual_workbook()
    {
        /*xlnt::workbook old_wb;
        saved_wb = save_virtual_workbook(old_wb);
        new_wb = load_workbook(BytesIO(saved_wb));
        assert new_wb;*/
    }

    void test_write_workbook_rels()
    {
        /*xlnt::workbook wb;
        content = write_workbook_rels(wb);
        assert_equals_file_content(os.path.join(DATADIR, "writer", "expected", \
            "workbook.xml.rels"), content);*/
    }

    void test_write_workbook()
    {
        /*xlnt::workbook wb;
        content = write_workbook(wb);
        assert_equals_file_content(os.path.join(DATADIR, "writer", "expected", \
            "workbook.xml"), content);*/
    }


    void test_write_string_table()
    {
        /*table = {"hello": 1, "world" : 2, "nice" : 3};
        content = write_string_table(table);
        assert_equals_file_content(os.path.join(DATADIR, "writer", "expected", \
            "sharedStrings.xml"), content);*/
    }

    void test_write_worksheet()
    {
        /*xlnt::workbook wb;
        auto ws = wb.create_sheet();
        ws.cell("F42") = "hello";
        content = write_worksheet(ws, {"hello": 0}, {});
        assert_equals_file_content(os.path.join(DATADIR, "writer", "expected", \
            "sheet1.xml"), content);*/
    }

    void test_write_hidden_worksheet()
    {
        /*xlnt::workbook wb;
        auto ws = wb.create_sheet();
        ws.sheet_state = ws.SHEETSTATE_HIDDEN;
        ws.cell("F42") = "hello";
        content = write_worksheet(ws, {"hello": 0}, {});
        assert_equals_file_content(os.path.join(DATADIR, "writer", "expected", \
            "sheet1.xml"), content);*/
    }

    void test_write_bool()
    {
        /*xlnt::workbook wb;
        auto ws = wb.create_sheet();
        ws.cell("F42") = False;
        ws.cell("F43") = True;
        content = write_worksheet(ws, {}, {});
        assert_equals_file_content(os.path.join(DATADIR, "writer", "expected", \
            "sheet1_bool.xml"), content);*/
    }

    void test_write_formula()
    {
        /*xlnt::workbook wb;
        auto ws = wb.create_sheet();
        ws.cell("F1") = 10;
        ws.cell("F2") = 32;
        ws.cell("F3") = "=F1+F2";
        content = write_worksheet(ws, {}, {});
        assert_equals_file_content(os.path.join(DATADIR, "writer", "expected", \
            "sheet1_formula.xml"), content);*/
    }

    void test_write_style()
    {
        /*xlnt::workbook wb;
        auto ws = wb.create_sheet();
        ws.cell("F1") = "13%";
        style_id_by_hash = StyleWriter(wb).get_style_by_hash();
        content = write_worksheet(ws, {}, style_id_by_hash);
        assert_equals_file_content(os.path.join(DATADIR, "writer", "expected", \
            "sheet1_style.xml"), content);*/
    }

    void test_write_height()
    {
        /*xlnt::workbook wb;
        auto ws = wb.create_sheet();
        ws.cell("F1") = 10;
        ws.row_dimensions[ws.cell("F1").row].height = 30;
        content = write_worksheet(ws, {}, {});
        assert_equals_file_content(os.path.join(DATADIR, "writer", "expected", \
            "sheet1_height.xml"), content);*/
    }

    void test_write_hyperlink()
    {
        /*xlnt::workbook wb;
        auto ws = wb.create_sheet();
        ws.cell("A1") = "test";
        ws.cell("A1").hyperlink = "http:test.com";
        content = write_worksheet(ws, {"test": 0}, {});
        assert_equals_file_content(os.path.join(DATADIR, "writer", "expected", \
            "sheet1_hyperlink.xml"), content);*/
    }

    void test_write_hyperlink_rels()
    {
        /*xlnt::workbook wb;
        auto ws = wb.create_sheet();
        TS_ASSERT_EQUALS(0, len(ws.relationships));
        ws.cell("A1") = "test";
        ws.cell("A1").hyperlink = "http:test.com/";
        TS_ASSERT_EQUALS(1, len(ws.relationships));
        ws.cell("A2") = "test";
        ws.cell("A2").hyperlink = "http:test2.com/";
        TS_ASSERT_EQUALS(2, len(ws.relationships));
        content = write_worksheet_rels(ws, 1);
        assert_equals_file_content(os.path.join(DATADIR, "writer", "expected", \
            "sheet1_hyperlink.xml.rels"), content);*/
    }

    void test_hyperlink_value()
    {
        xlnt::workbook wb;
        auto ws = wb.create_sheet();
        ws.get_cell("A1").set_hyperlink("http:test.com");
        TS_ASSERT_EQUALS("http:test.com", ws.get_cell("A1"));
        ws.get_cell("A1") = "test";
        TS_ASSERT_EQUALS("test", ws.get_cell("A1"));
    }

    void test_write_auto_filter()
    {
        /*xlnt::workbook wb;
        auto ws = wb[0];
        ws.cell("F42") = "hello";
        ws.auto_filter = "A1:F1";
        content = write_worksheet(ws, {"hello": 0}, {});
        assert_equals_file_content(os.path.join(DATADIR, "writer", "expected", \
            "sheet1_auto_filter.xml"), content);

        content = write_workbook(wb);
        assert_equals_file_content(os.path.join(DATADIR, "writer", "expected", \
            "workbook_auto_filter.xml"), content);*/
    }

    void test_freeze_panes_horiz()
    {
        /*xlnt::workbook wb;
        auto ws = wb.create_sheet();
        ws.cell("F42") = "hello";
        ws.freeze_panes = "A4";
        content = write_worksheet(ws, {"hello": 0}, {});
        assert_equals_file_content(os.path.join(DATADIR, "writer", "expected", \
            "sheet1_freeze_panes_horiz.xml"), content);*/
    }

    void test_freeze_panes_vert()
    {
        /*xlnt::workbook wb;
        auto ws = wb.create_sheet();
        ws.cell("F42") = "hello";
        ws.freeze_panes = "D1";
        content = write_worksheet(ws, {"hello": 0}, {});
        assert_equals_file_content(os.path.join(DATADIR, "writer", "expected", \
            "sheet1_freeze_panes_vert.xml"), content);*/
    }

    void test_freeze_panes_both()
    {
        /*xlnt::workbook wb;
        auto ws = wb.create_sheet();
        ws.cell("F42") = "hello";
        ws.freeze_panes = "D4";
        content = write_worksheet(ws, {"hello": 0}, {});
        assert_equals_file_content(os.path.join(DATADIR, "writer", "expected", \
            "sheet1_freeze_panes_both.xml"), content);*/
    }

    void test_long_number()
    {
        /*xlnt::workbook wb;
        auto ws = wb.create_sheet();
        ws.cell("A1") = 9781231231230;
        content = write_worksheet(ws, {}, {});
        assert_equals_file_content(os.path.join(DATADIR, "writer", "expected", \
            "long_number.xml"), content);*/
    }

    void test_decimal()
    {
        /*xlnt::workbook wb;
        auto ws = wb.create_sheet();
        ws.cell("A1") = decimal.Decimal("3.14");
        content = write_worksheet(ws, {}, {});
        assert_equals_file_content(os.path.join(DATADIR, "writer", "expected", \
            "decimal.xml"), content);*/
    }

    void test_short_number()
    {
        /*xlnt::workbook wb;
        auto ws = wb.create_sheet();
        ws.cell("A1") = 1234567890;
        content = write_worksheet(ws, {}, {});
        assert_equals_file_content(os.path.join(DATADIR, "writer", "expected", \
            "short_number.xml"), content);*/
    }
};
