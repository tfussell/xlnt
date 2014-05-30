#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include "pugixml.hpp"
#include <xlnt/xlnt.h>

class WorksheetTestSuite : public CxxTest::TestSuite
{
public:
    WorksheetTestSuite()
    {
    }

    void test_new_worksheet()
    {
        xlnt::worksheet ws = wb.create_sheet();
        TS_ASSERT(wb == ws.get_parent());
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
        auto cell = ws.get_cell("A1");
        TS_ASSERT_EQUALS(cell.get_reference().to_string(), "A1");
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
        ws.get_cell("A1") = "AAA";
        TS_ASSERT_EQUALS("A1", ws.calculate_dimension().to_string());
        ws.get_cell("B12") = "AAA";
        TS_ASSERT_EQUALS("A1:B12", ws.calculate_dimension().to_string());
    }

    void test_worksheet_range()
    {
        xlnt::worksheet ws(wb);
        auto xlrange = ws.get_range("A1:C4");
        TS_ASSERT_EQUALS(4, xlrange.num_rows());
        TS_ASSERT_EQUALS(3, xlrange[0].num_cells());
    }

    void test_worksheet_named_range()
    {
        xlnt::worksheet ws(wb);
        wb.create_named_range("test_range", ws, "C5");
        auto xlrange = ws.get_named_range("test_range");
        TS_ASSERT_EQUALS(1, xlrange.num_rows());
        TS_ASSERT_EQUALS(1, xlrange[0].num_cells());
        TS_ASSERT_EQUALS(5, xlrange[0][0].get_row());
    }

    void test_bad_named_range()
    {
        xlnt::worksheet ws(wb);
        TS_ASSERT_THROWS_ANYTHING(ws.get_range("bad_range"));
    }

    void test_named_range_wrong_sheet()
    {
        xlnt::worksheet ws1(wb);
        xlnt::worksheet ws2(wb);
        wb.create_named_range("wrong_sheet_range", ws1, "C5");
        TS_ASSERT_THROWS_ANYTHING(ws2.get_named_range("wrong_sheet_range"));
    }

    void test_cell_offset()
    {
        xlnt::worksheet ws(wb);
        TS_ASSERT_EQUALS("C17", ws.get_cell(xlnt::cell_reference("B15").make_offset(1, 2)).get_reference().to_string());
    }

    void test_range_offset()
    {
        xlnt::worksheet ws(wb);
        auto xlrange = ws.get_range(xlnt::range_reference("A1:C4").make_offset(3, 1));
        TS_ASSERT_EQUALS(4, xlrange.num_rows());
        TS_ASSERT_EQUALS(3, xlrange[0].num_cells());
        TS_ASSERT_EQUALS("D2", xlrange[0][0].get_reference().to_string());
    }

    void test_cell_alternate_coordinates()
    {
        xlnt::worksheet ws(wb);
        auto cell = ws.get_cell(xlnt::cell_reference(4, 8));
        TS_ASSERT_EQUALS("E9", cell.get_reference().to_string());
    }

    void test_cell_range_name()
    {
        xlnt::worksheet ws(wb);
        wb.create_named_range("test_range_single", ws, "B12");
        TS_ASSERT_THROWS(ws.get_cell("test_range_single"), xlnt::cell_coordinates_exception);
        auto c_range_name = ws.get_named_range("test_range_single");
        auto c_range_coord = ws.get_range("B12");
        auto c_cell = ws.get_cell("B12");
        TS_ASSERT_EQUALS(c_range_coord, c_range_name);
        TS_ASSERT(c_range_coord[0][0] == c_cell);
    }

    void test_garbage_collect()
    {
        xlnt::worksheet ws(wb);

        ws.get_cell("A1") = "";
        ws.get_cell("B2") = "0";
        ws.get_cell("C4") = 0;

        ws.garbage_collect();

        std::list<xlnt::cell> comparison_cells = {ws.get_cell("B2"), ws.get_cell("C4")};

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
        TS_SKIP("test_hyperlink_relationships");
        TS_ASSERT_EQUALS(ws.get_relationships().size(), 0);

        ws.get_cell("A1").set_hyperlink("http:test.com");
        TS_ASSERT_EQUALS(ws.get_relationships().size(), 1);
        TS_ASSERT_EQUALS("rId1", ws.get_cell("A1").get_hyperlink());
        TS_ASSERT_EQUALS("rId1", ws.get_relationships()[0].get_id());
        TS_ASSERT_EQUALS("http:test.com", ws.get_relationships()[0].get_target_uri());
        TS_ASSERT_EQUALS(xlnt::target_mode::external, ws.get_relationships()[0].get_target_mode());

        ws.get_cell("A2").set_hyperlink("http:test2.com");
        TS_ASSERT_EQUALS(ws.get_relationships().size(), 2);
        TS_ASSERT_EQUALS("rId2", ws.get_cell("A2").get_hyperlink());
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

        TS_ASSERT_EQUALS("This is A1", ws.get_cell("A1"));
        TS_ASSERT_EQUALS("This is B1", ws.get_cell("B1"));
    }

    void test_append_dict_letter()
    {
        xlnt::worksheet ws(wb);

        ws.append(std::unordered_map<std::string, std::string> {{"A", "This is A1"}, {"C", "This is C1"}});

        TS_ASSERT_EQUALS("This is A1", ws.get_cell("A1"));
        TS_ASSERT_EQUALS("This is C1", ws.get_cell("C1"));
    }

    void test_append_dict_index()
    {
        xlnt::worksheet ws(wb);

        ws.append(std::unordered_map<int, std::string> {{0, "This is A1"}, {2, "This is C1"}});

        TS_ASSERT_EQUALS("This is A1", ws.get_cell("A1"));
        TS_ASSERT_EQUALS("This is C1", ws.get_cell("C1"));
    }

    void test_append_2d_list()
    {
        xlnt::worksheet ws(wb);

        ws.append(std::vector<std::string> {"This is A1", "This is B1"});
        ws.append(std::vector<std::string> {"This is A2", "This is B2"});

        auto vals = ws.get_range("A1:B2");

        TS_ASSERT_EQUALS(vals[0][0], "This is A1");
        TS_ASSERT_EQUALS(vals[0][1], "This is B1");
        TS_ASSERT_EQUALS(vals[1][0], "This is A2");
        TS_ASSERT_EQUALS(vals[1][1], "This is B2");
    }

    void test_rows()
    {
        xlnt::worksheet ws(wb);

        ws.get_cell("A1") = "first";
        ws.get_cell("C9") = "last";

        auto rows = ws.rows();

        TS_ASSERT_EQUALS(rows.num_rows(), 9);

        TS_ASSERT_EQUALS(rows[0][0], "first");
        TS_ASSERT_EQUALS(rows[8][2], "last");
    }

    void test_auto_filter()
    {
        xlnt::worksheet ws(wb);

        ws.set_auto_filter(ws.get_range("a1:f1"));
        TS_ASSERT_EQUALS(ws.get_auto_filter(), "A1:F1");

        ws.unset_auto_filter();
        TS_ASSERT_EQUALS(ws.has_auto_filter(), false);

        ws.set_auto_filter("c1:g9");
        TS_ASSERT_EQUALS(ws.get_auto_filter(), "C1:G9");
    }

    void test_page_margins()
    {
        xlnt::worksheet ws(wb);

        ws.get_page_margins().set_left(2.0);
        ws.get_page_margins().set_right(2.0);
        ws.get_page_margins().set_top(2.0);
        ws.get_page_margins().set_bottom(2.0);
        ws.get_page_margins().set_header(1.5);
        ws.get_page_margins().set_footer(1.5);

        auto xml_string = xlnt::writer::write_worksheet(ws);
        pugi::xml_document doc;
        doc.load(xml_string.c_str());
        
        auto page_margins_node = doc.child("worksheet").child("pageMargins");
        TS_ASSERT_DIFFERS(page_margins_node.attribute("left"), nullptr);
        TS_ASSERT_EQUALS(page_margins_node.attribute("left").as_double(), 2.0);
        TS_ASSERT_DIFFERS(page_margins_node.attribute("right"), nullptr);
        TS_ASSERT_EQUALS(page_margins_node.attribute("right").as_double(), 2.0);
        TS_ASSERT_DIFFERS(page_margins_node.attribute("top"), nullptr);
        TS_ASSERT_EQUALS(page_margins_node.attribute("top").as_double(), 2.0);
        TS_ASSERT_DIFFERS(page_margins_node.attribute("bottom"), nullptr);
        TS_ASSERT_EQUALS(page_margins_node.attribute("bottom").as_double(), 2.0);
        TS_ASSERT_DIFFERS(page_margins_node.attribute("header"), nullptr);
        TS_ASSERT_EQUALS(page_margins_node.attribute("header").as_double(), 1.5);
        TS_ASSERT_DIFFERS(page_margins_node.attribute("footer"), nullptr);
        TS_ASSERT_EQUALS(page_margins_node.attribute("footer").as_double(), 1.5);

        xlnt::worksheet ws2(wb);
        xml_string = xlnt::writer::write_worksheet(ws2);
        doc.load(xml_string.c_str());
        TS_ASSERT_EQUALS(doc.child("worksheet").child("pageMargins"), nullptr);
    }

    void test_merge()
    {
        xlnt::worksheet ws(wb);

        std::vector<std::string> string_table = {"Cell A1", "Cell B1"};

        ws.get_cell("A1") = "Cell A1";
        ws.get_cell("B1") = "Cell B1";
        auto xml_string = xlnt::writer::write_worksheet(ws, string_table);
        pugi::xml_document doc;
        doc.load(xml_string.c_str());
        auto sheet_data_node = doc.child("worksheet").child("sheetData");
        auto row_node = sheet_data_node.find_child_by_attribute("row", "r", "1");
        auto cell_node = row_node.find_child_by_attribute("c", "r", "A1");
        TS_ASSERT_DIFFERS(cell_node, nullptr);
        TS_ASSERT_DIFFERS(cell_node.attribute("r"), nullptr);
        TS_ASSERT_EQUALS(std::string(cell_node.attribute("r").as_string()), "A1");
        TS_ASSERT_DIFFERS(cell_node.attribute("t"), nullptr);
        TS_ASSERT_EQUALS(std::string(cell_node.attribute("t").as_string()), "s");
        TS_ASSERT_DIFFERS(cell_node.child("v"), nullptr);
        TS_ASSERT_EQUALS(std::string(cell_node.child("v").text().as_string()), "0");
        cell_node = row_node.find_child_by_attribute("c", "r", "B1");
        TS_ASSERT_DIFFERS(cell_node, nullptr);
        TS_ASSERT_DIFFERS(cell_node.attribute("r"), nullptr);
        TS_ASSERT_EQUALS(std::string(cell_node.attribute("r").as_string()), "B1");
        TS_ASSERT_DIFFERS(cell_node.attribute("t"), nullptr);
        TS_ASSERT_EQUALS(std::string(cell_node.attribute("t").as_string()), "s");
        TS_ASSERT_DIFFERS(cell_node.child("v"), nullptr);
        TS_ASSERT_EQUALS(std::string(cell_node.child("v").text().as_string()), "1");

        ws.merge_cells("A1:B1");
        xml_string = xlnt::writer::write_worksheet(ws, string_table);
        doc.load(xml_string.c_str());
        sheet_data_node = doc.child("worksheet").child("sheetData");
        row_node = sheet_data_node.find_child_by_attribute("row", "r", "1");
        cell_node = row_node.find_child_by_attribute("c", "r", "A1");
        TS_ASSERT_DIFFERS(cell_node, nullptr);
        TS_ASSERT_DIFFERS(cell_node.attribute("r"), nullptr);
        TS_ASSERT_EQUALS(std::string(cell_node.attribute("r").as_string()), "A1");
        TS_ASSERT_DIFFERS(cell_node.attribute("t"), nullptr);
        TS_ASSERT_EQUALS(std::string(cell_node.attribute("t").as_string()), "s");
        TS_ASSERT_DIFFERS(cell_node.child("v"), nullptr);
        TS_ASSERT_EQUALS(std::string(cell_node.child("v").text().as_string()), "0");
        cell_node = row_node.find_child_by_attribute("c", "r", "B1");
        TS_ASSERT_DIFFERS(cell_node, nullptr);
        TS_ASSERT_EQUALS(cell_node.child("v"), nullptr);
        auto merge_node = doc.child("worksheet").child("mergeCells").child("mergeCell");
        TS_ASSERT_DIFFERS(merge_node, nullptr);
        TS_ASSERT_EQUALS(std::string(merge_node.attribute("ref").as_string()), "A1:B1");

        ws.unmerge_cells("A1:B1");
        xml_string = xlnt::writer::write_worksheet(ws, string_table);
        doc.load(xml_string.c_str());
        sheet_data_node = doc.child("worksheet").child("sheetData");
        row_node = sheet_data_node.find_child_by_attribute("row", "r", "1");
        cell_node = row_node.find_child_by_attribute("c", "r", "A1");
        TS_ASSERT_DIFFERS(cell_node, nullptr);
        TS_ASSERT_DIFFERS(cell_node.attribute("r"), nullptr);
        TS_ASSERT_EQUALS(std::string(cell_node.attribute("r").as_string()), "A1");
        TS_ASSERT_DIFFERS(cell_node.attribute("t"), nullptr);
        TS_ASSERT_EQUALS(std::string(cell_node.attribute("t").as_string()), "s");
        TS_ASSERT_DIFFERS(cell_node.child("v"), nullptr);
        TS_ASSERT_EQUALS(std::string(cell_node.child("v").text().as_string()), "0");
        cell_node = row_node.find_child_by_attribute("c", "r", "B1");
        TS_ASSERT_EQUALS(cell_node, nullptr);
        merge_node = doc.child("worksheet").child("mergeCells").child("mergeCell");
        TS_ASSERT_EQUALS(merge_node, nullptr);
    }

    void test_freeze()
    {
        xlnt::worksheet ws(wb);

        ws.freeze_panes(ws.get_cell("b2"));
        TS_ASSERT_EQUALS(ws.get_frozen_panes().to_string(), "B2");

        ws.unfreeze_panes();
        TS_ASSERT(!ws.has_frozen_panes());

        ws.freeze_panes("c5");
        TS_ASSERT_EQUALS(ws.get_frozen_panes().to_string(), "C5");

        ws.freeze_panes(ws.get_cell("A1"));
        TS_ASSERT(!ws.has_frozen_panes());
    }

    void test_printer_settings()
    {
        xlnt::worksheet ws(wb);

        ws.get_page_setup().set_orientation(xlnt::page_setup::orientation::landscape);
        ws.get_page_setup().set_paper_size(xlnt::page_setup::paper_size::tabloid);
        ws.get_page_setup().set_fit_to_page(true);
        ws.get_page_setup().set_fit_to_height(false);
        ws.get_page_setup().set_fit_to_width(true);

        auto xml_string = xlnt::writer::write_worksheet(ws);

        pugi::xml_document doc;
        doc.load(xml_string.c_str());

        auto page_setup_node = doc.child("worksheet").child("pageSetup");
        TS_ASSERT_DIFFERS(page_setup_node, nullptr);
        TS_ASSERT_DIFFERS(page_setup_node.attribute("orientation"), nullptr);
        TS_ASSERT_EQUALS(std::string(page_setup_node.attribute("orientation").as_string()), "landscape");
        TS_ASSERT_DIFFERS(page_setup_node.attribute("paperSize"), nullptr);
        TS_ASSERT_EQUALS(page_setup_node.attribute("paperSize").as_int(), 3);
        TS_ASSERT_DIFFERS(page_setup_node.attribute("fitToHeight"), nullptr);
        TS_ASSERT_EQUALS(page_setup_node.attribute("fitToHeight").as_int(), 0);
        TS_ASSERT_DIFFERS(page_setup_node.attribute("fitToWidth"), nullptr);
        TS_ASSERT_EQUALS(page_setup_node.attribute("fitToWidth").as_int(), 1);
        TS_ASSERT_DIFFERS(doc.child("worksheet").child("pageSetUpPr").attribute("fitToPage"), nullptr);
        TS_ASSERT_EQUALS(doc.child("worksheet").child("pageSetUpPr").attribute("fitToPage").as_int(), 1);

        xlnt::worksheet ws2(wb);
        xml_string = xlnt::writer::write_worksheet(ws2);
        doc.load(xml_string.c_str());
        TS_ASSERT_EQUALS(doc.child("worksheet").child("pageSetup"), nullptr);
        TS_ASSERT_EQUALS(doc.child("worksheet").child("pageSetUpPr"), nullptr);
    }

private:
    xlnt::workbook wb;
};
