#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include "pugixml.hpp"
#include "../xlnt.h"

class WorksheetTestSuite : public CxxTest::TestSuite
{
public:
    WorksheetTestSuite()
    {
    }

    void test_new_worksheet()
    {
        xlnt::worksheet ws = wb.create_sheet();
        TS_ASSERT_EQUALS(&wb, &ws.get_parent());
    }

    void test_new_sheet_name()
    {
        xlnt::workbook wb;
        wb.remove_sheet(wb.get_active_sheet());
        xlnt::worksheet ws = wb.create_sheet();
        TS_ASSERT_EQUALS(ws.to_string(), "<Worksheet \"Sheet1\">");
    }

    void test_get_cell()
    {
        xlnt::worksheet ws(wb);
        auto cell = ws.cell("A1");
        TS_ASSERT_EQUALS(cell.get_coordinate(), "A1");
    }

    void test_set_bad_title()
    {
        std::string title(50, 'X');
        TS_ASSERT_THROWS(wb.create_sheet(title), xlnt::bad_sheet_title);
    }

    void test_set_bad_title_character()
    {
        TS_ASSERT_THROWS(wb.create_sheet("["), xlnt::bad_sheet_title);
        TS_ASSERT_THROWS(wb.create_sheet("]"), xlnt::bad_sheet_title);
        TS_ASSERT_THROWS(wb.create_sheet("*"), xlnt::bad_sheet_title);
        TS_ASSERT_THROWS(wb.create_sheet(":"), xlnt::bad_sheet_title);
        TS_ASSERT_THROWS(wb.create_sheet("?"), xlnt::bad_sheet_title);
        TS_ASSERT_THROWS(wb.create_sheet("/"), xlnt::bad_sheet_title);
        TS_ASSERT_THROWS(wb.create_sheet("\\"), xlnt::bad_sheet_title);
    }

    void test_worksheet_dimension()
    {
        xlnt::worksheet ws(wb);
        TS_ASSERT_EQUALS("A1:A1", ws.calculate_dimension());
        ws.cell("B12") = "AAA";
        TS_ASSERT_EQUALS("A1:B12", ws.calculate_dimension());
    }

    void test_worksheet_range()
    {
        xlnt::worksheet ws(wb);
        auto xlrange = ws.range("A1:C4");
        TS_ASSERT_EQUALS(4, xlrange.size());
        TS_ASSERT_EQUALS(3, xlrange[0].size());
    }

    void test_worksheet_named_range()
    {
        xlnt::worksheet ws(wb);
        wb.create_named_range("test_range", ws, "C5");
        auto xlrange = ws.range("test_range");
        TS_ASSERT_EQUALS(1, xlrange.size());
        TS_ASSERT_EQUALS(1, xlrange[0].size());
        TS_ASSERT_EQUALS(5, xlrange[0][0].get_row());
    }

    void test_bad_named_range()
    {
        xlnt::worksheet ws(wb);
        TS_ASSERT_THROWS_ANYTHING(ws.range("bad_range"));
    }

    void test_named_range_wrong_sheet()
    {
        xlnt::worksheet ws1(wb);
        xlnt::worksheet ws2(wb);
        wb.create_named_range("wrong_sheet_range", ws1, "C5");
        TS_ASSERT_THROWS_ANYTHING(ws2.range("wrong_sheet_range"));
    }

    void test_cell_offset()
    {
        xlnt::worksheet ws(wb);
        TS_ASSERT_EQUALS("C17", ws.cell("B15").get_offset(2, 1).get_coordinate());
    }

    void test_range_offset()
    {
        xlnt::worksheet ws(wb);
        auto xlrange = ws.range("A1:C4", 1, 3);
        TS_ASSERT_EQUALS(4, xlrange.size());
        TS_ASSERT_EQUALS(3, xlrange[0].size());
        TS_ASSERT_EQUALS("D2", xlrange[0][0].get_coordinate());
    }

    void test_cell_alternate_coordinates()
    {
        xlnt::worksheet ws(wb);
        auto cell = ws.cell(4, 8);
        TS_ASSERT_EQUALS("E9", cell.get_coordinate());
    }

    void test_cell_range_name()
    {
        xlnt::worksheet ws(wb);
        wb.create_named_range("test_range_single", ws, "B12");
        TS_ASSERT_THROWS(ws.cell("test_range_single"), xlnt::bad_cell_coordinates);
        auto c_range_name = ws.range("test_range_single");
        auto c_range_coord = ws.range("B12");
        auto c_cell = ws.cell("B12");
        TS_ASSERT_EQUALS(c_range_coord, c_range_name);
        TS_ASSERT(c_range_coord[0][0] == c_cell);
    }

    void test_garbage_collect()
    {
        xlnt::worksheet ws(wb);

        ws.cell("A1") = "";
        ws.cell("B2") = "0";
        ws.cell("C4") = 0;

        ws.garbage_collect();

        std::list<xlnt::cell> comparison_cells = {ws.cell("B2"), ws.cell("C4")};

        for(auto cell : ws.get_cell_collection())
        {
            auto match = std::find(comparison_cells.begin(), comparison_cells.end(), cell);
            TS_ASSERT_DIFFERS(match, comparison_cells.end());
            comparison_cells.erase(match);
        }

        TS_ASSERT(comparison_cells.empty());
    }

    void test_hyperlink_relationships()
    {
        xlnt::worksheet ws(wb);

        TS_ASSERT_EQUALS(ws.get_relationships().size(), 0);

        ws.cell("A1").set_hyperlink("http:test.com");
        TS_ASSERT_EQUALS(ws.get_relationships().size(), 1);
        TS_ASSERT_EQUALS("rId1", ws.cell("A1").get_hyperlink_rel_id());
        TS_ASSERT_EQUALS("rId1", ws.get_relationships()[0].get_id());
        TS_ASSERT_EQUALS("http:test.com", ws.get_relationships()[0].get_target_uri());
        TS_ASSERT_EQUALS(xlnt::target_mode::external, ws.get_relationships()[0].get_target_mode());

        ws.cell("A2").set_hyperlink("http:test2.com");
        TS_ASSERT_EQUALS(ws.get_relationships().size(), 2);
        TS_ASSERT_EQUALS("rId2", ws.cell("A2").get_hyperlink_rel_id());
        TS_ASSERT_EQUALS("rId2", ws.get_relationships()[1].get_id());
        TS_ASSERT_EQUALS("http:test2.com", ws.get_relationships()[1].get_target_uri());
        TS_ASSERT_EQUALS(xlnt::target_mode::external, ws.get_relationships()[1].get_target_mode());
    }

    void test_bad_relationship_type()
    {
        xlnt::relationship rel("bad_type");
    }

    void test_append_list()
    {
        xlnt::worksheet ws(wb);

        ws.append(std::vector<std::string> {"This is A1", "This is B1"});

        TS_ASSERT_EQUALS("This is A1", ws.cell("A1"));
        TS_ASSERT_EQUALS("This is B1", ws.cell("B1"));
    }

    void test_append_dict_letter()
    {
        xlnt::worksheet ws(wb);

        ws.append(std::unordered_map<std::string, std::string> {{"A", "This is A1"}, {"C", "This is C1"}});

        TS_ASSERT_EQUALS("This is A1", ws.cell("A1"));
        TS_ASSERT_EQUALS("This is C1", ws.cell("C1"));
    }

    void test_append_dict_index()
    {
        xlnt::worksheet ws(wb);

        ws.append(std::unordered_map<int, std::string> {{0, "This is A1"}, {2, "This is C1"}});

        TS_ASSERT_EQUALS("This is A1", ws.cell("A1"));
        TS_ASSERT_EQUALS("This is C1", ws.cell("C1"));
    }

    void test_append_2d_list()
    {
        xlnt::worksheet ws(wb);

        ws.append(std::vector<std::string> {"This is A1", "This is B1"});
        ws.append(std::vector<std::string> {"This is A2", "This is B2"});

        auto vals = ws.range("A1:B2");

        TS_ASSERT_EQUALS(vals[0][0], "This is A1");
        TS_ASSERT_EQUALS(vals[0][1], "This is B1");
        TS_ASSERT_EQUALS(vals[1][0], "This is A2");
        TS_ASSERT_EQUALS(vals[1][1], "This is B2");
    }

    void test_rows()
    {
        xlnt::worksheet ws(wb);

        ws.cell("A1") = "first";
        ws.cell("C9") = "last";

        auto rows = ws.rows();

        TS_ASSERT_EQUALS(rows.size(), 9);

        TS_ASSERT_EQUALS(rows[0][0], "first");
        TS_ASSERT_EQUALS(rows[8][2], "last");
    }

    void test_cols()
    {
        xlnt::worksheet ws(wb);

        ws.cell("A1") = "first";
        ws.cell("C9") = "last";

        auto cols = ws.columns();

        TS_ASSERT_EQUALS(cols.size(), 3);

        TS_ASSERT_EQUALS(cols[0][0], "first");
        TS_ASSERT_EQUALS(cols[2][8], "last");
    }

    void test_auto_filter()
    {
        xlnt::worksheet ws(wb);

        /*ws.set_auto_filter(ws.range("a1:f1"));
        TS_ASSERT_EQUALS(ws.get_auto_filter(), "A1:F1");

        ws.set_auto_filter("");
        assert ws.auto_filter is None;

        ws.auto_filter = "c1:g9";
        assert ws.auto_filter == "C1:G9";*/
    }

    void test_page_margins()
    {
        xlnt::worksheet ws(wb);

        /*ws.get_page_margins().set_left(2.0);
        ws.get_page_margins().set_right(2.0);
        ws.get_page_margins().set_top(2.0);
        ws.get_page_margins().set_bottom(2.0);
        ws.get_page_margins().set_header(1.5);
        ws.get_page_margins().set_footer(1.5);*/

        //auto xml_string = xlnt::writer::write_worksheet(ws);

        //assert "<pageMargins left="2.00" right="2.00" top="2.00" bottom="2.00" header="1.50" footer="1.50"></pageMargins>" in xml_string;

        xlnt::worksheet ws2(wb);
        //xml_string = xlnt::writer::write_worksheet(ws2);
        //assert "<pageMargins" not in xml_string;
    }

    void test_merge()
    {
        xlnt::worksheet ws(wb);

        std::unordered_map<std::string, std::string> string_table = {{"", ""}, {"Cell A1", "Cell A1"}, {"Cell B1", "Cell B1"}};

        ws.cell("A1") = "Cell A1";
        ws.cell("B1") = "Cell B1";
        //auto xml_string = xlnt::writer::write_worksheet(ws);// , string_table);
        /*assert "<c r="B1" t="s"><v>Cell B1</v></c>" in xml_string;

        ws.merge_cells("A1:B1");
        xml_string = xlnt::writer::write_worksheet(ws, string_table, None);
        assert "<c r="B1" t="s"><v>Cell B1</v></c>" not in xml_string;
        assert "<mergeCells><mergeCell ref="A1:B1"></mergeCell></mergeCells>" in xml_string;

        ws.unmerge_cells("A1:B1");
        xml_string = xlnt::writer::write_worksheet(ws, string_table, None);
        assert "<mergeCell ref="A1:B1"></mergeCell>" not in xml_string;*/
    }

    void test_freeze()
    {
        xlnt::worksheet ws(wb);

        ws.set_freeze_panes(ws.cell("b2"));
        TS_ASSERT_EQUALS(ws.get_freeze_panes().get_address(), "B2");

        ws.unfreeze_panes();
        TS_ASSERT_EQUALS(ws.get_freeze_panes(), nullptr);

        ws.set_freeze_panes("c5");
        TS_ASSERT_EQUALS(ws.get_freeze_panes().get_address(), "C5");

        ws.set_freeze_panes(ws.cell("A1"));
        TS_ASSERT_EQUALS(ws.get_freeze_panes(), nullptr);
    }

    void test_printer_settings()
    {
        xlnt::worksheet ws(wb);

        ws.get_page_setup().orientation = xlnt::page_setup::Orientation::Landscape;
        ws.get_page_setup().paper_size = xlnt::page_setup::PaperSize::Tabloid;
        ws.get_page_setup().fit_to_page = true;
        ws.get_page_setup().fit_to_height = false;
        ws.get_page_setup().fit_to_width = true;
        //auto xml_string = xlnt::writer::write_worksheet(ws);
        //pugi::xml_document doc;
        //TS_ASSERT("<pageSetup orientation=\"landscape\" paperSize=\"3\" fitToHeight=\"0\" fitToWidth=\"1\"></pageSetup>" in xml_string);
        //TS_ASSERT("<pageSetUpPr fitToPage=\"1\"></pageSetUpPr>" in xml_string);

        xlnt::worksheet ws2(wb);
        //xml_string = xlnt::writer::write_worksheet(ws2);
        //TS_ASSERT("<pageSetup" not in xml_string);
        //TS_ASSERT("<pageSetUpPr" not in xml_string);
    }

private:
    xlnt::workbook wb;
};
