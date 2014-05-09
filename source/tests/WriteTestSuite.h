#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include "../xlnt.h"

class WriteTestSuite : public CxxTest::TestSuite
{
public:
    WriteTestSuite()
    {

    }

    void test_write_empty_workbook()
    {
        //make_tmpdir();
        //wb = Workbook();
        //dest_filename = os.path.join(TMPDIR, "empty_book.xlsx");
        //save_workbook(wb, dest_filename);
        //assert os.path.isfile(dest_filename);
        //clean_tmpdir();
    }

    void test_write_virtual_workbook()
    {
        //old_wb = Workbook();
        //saved_wb = save_virtual_workbook(old_wb);
        //new_wb = load_workbook(BytesIO(saved_wb));
        //assert new_wb;
    }

    void test_write_workbook_rels()
    {
        //wb = Workbook();
        //content = write_workbook_rels(wb);
        //assert_equals_file_content(os.path.join(DATADIR, "writer", "expected", \
        //    "workbook.xml.rels"), content);
    }

    void test_write_workbook()
    {
        //wb = Workbook();
        //content = write_workbook(wb);
        //assert_equals_file_content(os.path.join(DATADIR, "writer", "expected", \
        //    "workbook.xml"), content);
    }


    void test_write_string_table()
    {
        //table = {"hello": 1, "world" : 2, "nice" : 3};
        //content = write_string_table(table);
        //assert_equals_file_content(os.path.join(DATADIR, "writer", "expected", \
        //    "sharedStrings.xml"), content);
    }

    void test_write_worksheet()
    {
        //wb = Workbook();
        //ws = wb.create_sheet();
        //ws.cell("F42").value = "hello";
        //content = write_worksheet(ws, {"hello": 0}, {});
        //assert_equals_file_content(os.path.join(DATADIR, "writer", "expected", \
        //    "sheet1.xml"), content);
    }

    void test_write_hidden_worksheet()
    {
        //wb = Workbook();
        //ws = wb.create_sheet();
        //ws.sheet_state = ws.SHEETSTATE_HIDDEN;
        //ws.cell("F42").value = "hello";
        //content = write_worksheet(ws, {"hello": 0}, {});
        //assert_equals_file_content(os.path.join(DATADIR, "writer", "expected", \
        //    "sheet1.xml"), content);
    }

    void test_write_bool()
    {
        //wb = Workbook();
        //ws = wb.create_sheet();
        //ws.cell("F42").value = False;
        //ws.cell("F43").value = True;
        //content = write_worksheet(ws, {}, {});
        //assert_equals_file_content(os.path.join(DATADIR, "writer", "expected", \
        //    "sheet1_bool.xml"), content);
    }

    void test_write_formula()
    {
        //wb = Workbook();
        //ws = wb.create_sheet();
        //ws.cell("F1").value = 10;
        //ws.cell("F2").value = 32;
        //ws.cell("F3").value = "=F1+F2";
        //content = write_worksheet(ws, {}, {});
        //assert_equals_file_content(os.path.join(DATADIR, "writer", "expected", \
        //    "sheet1_formula.xml"), content);
    }

    void test_write_style()
    {
        //wb = Workbook();
        //ws = wb.create_sheet();
        //ws.cell("F1").value = "13%";
        //style_id_by_hash = StyleWriter(wb).get_style_by_hash();
        //content = write_worksheet(ws, {}, style_id_by_hash);
        //assert_equals_file_content(os.path.join(DATADIR, "writer", "expected", \
        //    "sheet1_style.xml"), content);
    }

    void test_write_height()
    {
        //wb = Workbook();
        //ws = wb.create_sheet();
        //ws.cell("F1").value = 10;
        //ws.row_dimensions[ws.cell("F1").row].height = 30;
        //content = write_worksheet(ws, {}, {});
        //assert_equals_file_content(os.path.join(DATADIR, "writer", "expected", \
        //    "sheet1_height.xml"), content);
    }

    void test_write_hyperlink()
    {
        //wb = Workbook();
        //ws = wb.create_sheet();
        //ws.cell("A1").value = "test";
        //ws.cell("A1").hyperlink = "http://test.com";
        //content = write_worksheet(ws, {"test": 0}, {});
        //assert_equals_file_content(os.path.join(DATADIR, "writer", "expected", \
        //    "sheet1_hyperlink.xml"), content);
    }

    void test_write_hyperlink_rels()
    {
        //wb = Workbook();
        //ws = wb.create_sheet();
        //TS_ASSERT_EQUALS(0, len(ws.relationships));
        //ws.cell("A1").value = "test";
        //ws.cell("A1").hyperlink = "http://test.com/";
        //TS_ASSERT_EQUALS(1, len(ws.relationships));
        //ws.cell("A2").value = "test";
        //ws.cell("A2").hyperlink = "http://test2.com/";
        //TS_ASSERT_EQUALS(2, len(ws.relationships));
        //content = write_worksheet_rels(ws, 1);
        //assert_equals_file_content(os.path.join(DATADIR, "writer", "expected", \
        //    "sheet1_hyperlink.xml.rels"), content);
    }

    void test_hyperlink_value()
    {
        //wb = Workbook();
        //ws = wb.create_sheet();
        //ws.cell("A1").hyperlink = "http://test.com";
        //TS_ASSERT_EQUALS("http://test.com", ws.cell("A1").value);
        //ws.cell("A1").value = "test";
        //TS_ASSERT_EQUALS("test", ws.cell("A1").value);
    }

    void test_write_auto_filter()
    {
        //wb = Workbook();
        //ws = wb.worksheets[0];
        //ws.cell("F42").value = "hello";
        //ws.auto_filter = "A1:F1";
        //content = write_worksheet(ws, {"hello": 0}, {});
        //assert_equals_file_content(os.path.join(DATADIR, "writer", "expected", \
        //    "sheet1_auto_filter.xml"), content);

        //content = write_workbook(wb);
        //assert_equals_file_content(os.path.join(DATADIR, "writer", "expected", \
        //    "workbook_auto_filter.xml"), content);
    }

    void test_freeze_panes_horiz()
    {
        //wb = Workbook();
        //ws = wb.create_sheet();
        //ws.cell("F42").value = "hello";
        //ws.freeze_panes = "A4";
        //content = write_worksheet(ws, {"hello": 0}, {});
        //assert_equals_file_content(os.path.join(DATADIR, "writer", "expected", \
        //    "sheet1_freeze_panes_horiz.xml"), content);
    }

    void test_freeze_panes_vert()
    {
        //wb = Workbook();
        //ws = wb.create_sheet();
        //ws.cell("F42").value = "hello";
        //ws.freeze_panes = "D1";
        //content = write_worksheet(ws, {"hello": 0}, {});
        //assert_equals_file_content(os.path.join(DATADIR, "writer", "expected", \
        //    "sheet1_freeze_panes_vert.xml"), content);
    }

    void test_freeze_panes_both()
    {
        //wb = Workbook();
        //ws = wb.create_sheet();
        //ws.cell("F42").value = "hello";
        //ws.freeze_panes = "D4";
        //content = write_worksheet(ws, {"hello": 0}, {});
        //assert_equals_file_content(os.path.join(DATADIR, "writer", "expected", \
        //    "sheet1_freeze_panes_both.xml"), content);
    }

    void test_long_number()
    {
        //wb = Workbook();
        //ws = wb.create_sheet();
        //ws.cell("A1").value = 9781231231230;
        //content = write_worksheet(ws, {}, {});
        //assert_equals_file_content(os.path.join(DATADIR, "writer", "expected", \
        //    "long_number.xml"), content);
    }

    void test_decimal()
    {
        //wb = Workbook();
        //ws = wb.create_sheet();
        //ws.cell("A1").value = decimal.Decimal("3.14");
        //content = write_worksheet(ws, {}, {});
        //assert_equals_file_content(os.path.join(DATADIR, "writer", "expected", \
        //    "decimal.xml"), content);
    }

    void test_short_number()
    {
        //wb = Workbook();
        //ws = wb.create_sheet();
        //ws.cell("A1").value = 1234567890;
        //content = write_worksheet(ws, {}, {});
        //assert_equals_file_content(os.path.join(DATADIR, "writer", "expected", \
        //    "short_number.xml"), content);
    }
};
