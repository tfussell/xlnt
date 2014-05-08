#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/workbook.h>
#include <xlnt/worksheet.h>

class WorksheetTestSuite : public CxxTest::TestSuite
{
public:
    WorksheetTestSuite()
    {
    }/*
        cls.wb = Workbook()
    }
    */

    void test_new_worksheet()
    {
        //ws = Worksheet(wb);
        //TS_ASSERT_EQUALS(wb, ws._parent);
    }

    void test_new_sheet_name()
    {
        //wb.worksheets = [];
        //ws = Worksheet(wb, title = "");
        //TS_ASSERT_EQUALS(repr(ws), "<Worksheet "Sheet1">");
    }

    void test_get_cell()
    {
        //ws = Worksheet(wb);
        //cell = ws.cell("A1");
        //TS_ASSERT_EQUALS(cell.get_coordinate(), "A1");
    }

    void test_set_bad_title()
    {
        //Worksheet(wb, "X" * 50);
    }

    void test_set_bad_title_character()
    {
        //assert_raises(SheetTitleException, Worksheet, wb, "[");
        //assert_raises(SheetTitleException, Worksheet, wb, "]");
        //assert_raises(SheetTitleException, Worksheet, wb, "*");
        //assert_raises(SheetTitleException, Worksheet, wb, ":");
        //assert_raises(SheetTitleException, Worksheet, wb, "?");
        //assert_raises(SheetTitleException, Worksheet, wb, "/");
        //assert_raises(SheetTitleException, Worksheet, wb, "\\");
    }

    void test_worksheet_dimension()
    {
        //ws = Worksheet(wb);
        //TS_ASSERT_EQUALS("A1:A1", ws.calculate_dimension());
        //ws.cell("B12").value = "AAA";
        //TS_ASSERT_EQUALS("A1:B12", ws.calculate_dimension());
    }

    void test_worksheet_range()
    {
        //ws = Worksheet(wb);
        //xlrange = ws.range("A1:C4");
        //assert isinstance(xlrange, tuple);
        //TS_ASSERT_EQUALS(4, len(xlrange));
        //TS_ASSERT_EQUALS(3, len(xlrange[0]));
    }

    void test_worksheet_named_range()
    {
        //ws = Worksheet(wb);
        //wb.create_named_range("test_range", ws, "C5");
        //xlrange = ws.range("test_range");
        //assert isinstance(xlrange, Cell);
        //TS_ASSERT_EQUALS(5, xlrange.row);
    }

    void test_bad_named_range()
    {
        //ws = Worksheet(wb);
        //ws.range("bad_range");
    }

    void test_named_range_wrong_sheet()
    {
        //ws1 = Worksheet(wb);
        //ws2 = Worksheet(wb);
        //wb.create_named_range("wrong_sheet_range", ws1, "C5");
        //ws2.range("wrong_sheet_range");
    }

    void test_cell_offset()
    {
        //ws = Worksheet(wb);
        //TS_ASSERT_EQUALS("C17", ws.cell("B15").offset(2, 1).get_coordinate());
    }

    void test_range_offset()
    {
        //ws = Worksheet(wb);
        //xlrange = ws.range("A1:C4", 1, 3);
        //assert isinstance(xlrange, tuple);
        //TS_ASSERT_EQUALS(4, len(xlrange));
        //TS_ASSERT_EQUALS(3, len(xlrange[0]));
        //TS_ASSERT_EQUALS("D2", xlrange[0][0].get_coordinate());
    }

    void test_cell_alternate_coordinates()
    {
        //ws = Worksheet(wb);
        //cell = ws.cell(row = 8, column = 4);
        //TS_ASSERT_EQUALS("E9", cell.get_coordinate());
    }

    void test_cell_insufficient_coordinates()
    {
        //ws = Worksheet(wb);
        //cell = ws.cell(row = 8);
    }

    void test_cell_range_name()
    {
        //ws = Worksheet(wb);
        //wb.create_named_range("test_range_single", ws, "B12");
        //assert_raises(CellCoordinatesException, ws.cell, "test_range_single");
        //c_range_name = ws.range("test_range_single");
        //c_range_coord = ws.range("B12");
        //c_cell = ws.cell("B12");
        //TS_ASSERT_EQUALS(c_range_coord, c_range_name);
        //TS_ASSERT_EQUALS(c_range_coord, c_cell);
    }

    void test_garbage_collect()
    {
        //ws = Worksheet(wb);
        //ws.cell("A1").value = "";
        //ws.cell("B2").value = "0";
        //ws.cell("C4").value = 0;
        //ws.garbage_collect();
        //for i, cell in enumerate(ws.get_cell_collection())
        //{
        //    TS_ASSERT_EQUALS(cell, [ws.cell("B2"), ws.cell("C4")][i]);
        //}
    }

    void test_hyperlink_relationships()
    {
        //ws = Worksheet(wb);
        //TS_ASSERT_EQUALS(len(ws.relationships), 0);

        //ws.cell("A1").hyperlink = "http://test.com";
        //TS_ASSERT_EQUALS(len(ws.relationships), 1);
        //TS_ASSERT_EQUALS("rId1", ws.cell("A1").hyperlink_rel_id);
        //TS_ASSERT_EQUALS("rId1", ws.relationships[0].id);
        //TS_ASSERT_EQUALS("http://test.com", ws.relationships[0].target);
        //TS_ASSERT_EQUALS("External", ws.relationships[0].target_mode);

        //ws.cell("A2").hyperlink = "http://test2.com";
        //TS_ASSERT_EQUALS(len(ws.relationships), 2);
        //TS_ASSERT_EQUALS("rId2", ws.cell("A2").hyperlink_rel_id);
        //TS_ASSERT_EQUALS("rId2", ws.relationships[1].id);
        //TS_ASSERT_EQUALS("http://test2.com", ws.relationships[1].target);
        //TS_ASSERT_EQUALS("External", ws.relationships[1].target_mode);
    }

    void test_bad_relationship_type()
    {
        //rel = Relationship("bad_type");
    }

    void test_append_list()
    {
        //ws = Worksheet(wb);

        //ws.append(["This is A1", "This is B1"]);

        //TS_ASSERT_EQUALS("This is A1", ws.cell("A1").value);
        //TS_ASSERT_EQUALS("This is B1", ws.cell("B1").value);
    }

    void test_append_dict_letter()
    {
        //ws = Worksheet(wb);

        //ws.append({"A" : "This is A1", "C" : "This is C1"});

        //TS_ASSERT_EQUALS("This is A1", ws.cell("A1").value);
        //TS_ASSERT_EQUALS("This is C1", ws.cell("C1").value);
    }

    void test_append_dict_index()
    {
        //ws = Worksheet(wb);

        //ws.append({0 : "This is A1", 2 : "This is C1"});

        //TS_ASSERT_EQUALS("This is A1", ws.cell("A1").value);
        //TS_ASSERT_EQUALS("This is C1", ws.cell("C1").value);
    }

    void test_bad_append()
    {
        //ws = Worksheet(wb);
        //ws.append("test");
    }

    void test_append_2d_list()
    {
        //ws = Worksheet(wb);

        //ws.append(["This is A1", "This is B1"]);
        //ws.append(["This is A2", "This is B2"]);

        //vals = ws.range("A1:B2");

        //TS_ASSERT_EQUALS((("This is A1", "This is B1"),
        //    ("This is A2", "This is B2"), ), flatten(vals));
    }

    void test_rows()
    {
        //ws = Worksheet(wb);

        //ws.cell("A1").value = "first";
        //ws.cell("C9").value = "last";

        //rows = ws.rows;

        //TS_ASSERT_EQUALS(len(rows), 9);

        //TS_ASSERT_EQUALS(rows[0][0].value, "first");
        //TS_ASSERT_EQUALS(rows[-1][-1].value, "last");
    }

    void test_cols()
    {
        //ws = Worksheet(wb);

        //ws.cell("A1").value = "first";
        //ws.cell("C9").value = "last";

        //cols = ws.columns;

        //TS_ASSERT_EQUALS(len(cols), 3);

        //TS_ASSERT_EQUALS(cols[0][0].value, "first");
        //TS_ASSERT_EQUALS(cols[-1][-1].value, "last");
    }

    void test_auto_filter()
    {
        //ws = Worksheet(wb);
        //ws.auto_filter = ws.range("a1:f1");
        //assert ws.auto_filter == "A1:F1";

        //ws.auto_filter = "";
        //assert ws.auto_filter is None;

        //ws.auto_filter = "c1:g9";
        //assert ws.auto_filter == "C1:G9";
    }

    void test_page_margins()
    {
        //ws = Worksheet(wb);
        //ws.page_margins.left = 2.0;
        //ws.page_margins.right = 2.0;
        //ws.page_margins.top = 2.0;
        //ws.page_margins.bottom = 2.0;
        //ws.page_margins.header = 1.5;
        //ws.page_margins.footer = 1.5;
        //xml_string = write_worksheet(ws, None, None);
        //assert "<pageMargins left="2.00" right="2.00" top="2.00" bottom="2.00" header="1.50" footer="1.50"></pageMargins>" in xml_string;

        //ws = Worksheet(wb);
        //xml_string = write_worksheet(ws, None, None);
        //assert "<pageMargins" not in xml_string;
    }

    void test_merge()
    {
        //ws = Worksheet(wb);
        //string_table = {"":"", "Cell A1" : "Cell A1", "Cell B1" : "Cell B1"};

        //ws.cell("A1").value = "Cell A1";
        //ws.cell("B1").value = "Cell B1";
        //xml_string = write_worksheet(ws, string_table, None);
        //assert "<c r="B1" t="s"><v>Cell B1</v></c>" in xml_string;

        //ws.merge_cells("A1:B1");
        //xml_string = write_worksheet(ws, string_table, None);
        //assert "<c r="B1" t="s"><v>Cell B1</v></c>" not in xml_string;
        //assert "<mergeCells><mergeCell ref="A1:B1"></mergeCell></mergeCells>" in xml_string;

        //ws.unmerge_cells("A1:B1");
        //xml_string = write_worksheet(ws, string_table, None);
        //assert "<mergeCell ref="A1:B1"></mergeCell>" not in xml_string;
    }

    void test_freeze()
    {
        //ws = Worksheet(wb);
        //ws.freeze_panes = ws.cell("b2");
        //assert ws.freeze_panes == "B2";

        //ws.freeze_panes = "";
        //assert ws.freeze_panes is None;

        //ws.freeze_panes = "c5";
        //assert ws.freeze_panes == "C5";

        //ws.freeze_panes = ws.cell("A1");
        //assert ws.freeze_panes is None;
    }

    void test_printer_settings()
    {
        //ws = Worksheet(wb);
        //ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE;
        //ws.page_setup.paperSize = ws.PAPERSIZE_TABLOID;
        //ws.page_setup.fitToPage = True;
        //ws.page_setup.fitToHeight = 0;
        //ws.page_setup.fitToWidth = 1;
        //xml_string = write_worksheet(ws, None, None);
        //assert "<pageSetup orientation="landscape" paperSize="3" fitToHeight="0" fitToWidth="1"></pageSetup>" in xml_string;
        //assert "<pageSetUpPr fitToPage="1"></pageSetUpPr>" in xml_string;

        //ws = Worksheet(wb);
        //xml_string = write_worksheet(ws, None, None);
        //assert "<pageSetup" not in xml_string;
        //assert "<pageSetUpPr" not in xml_string;
    }
};
