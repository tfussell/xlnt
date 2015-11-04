#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/serialization/worksheet_serializer.hpp>
#include <xlnt/worksheet/worksheet.hpp>

class test_worksheet : public CxxTest::TestSuite
{
public:
    void test_new_worksheet()
    {
        xlnt::worksheet ws = wb_.create_sheet();
        TS_ASSERT(wb_ == ws.get_parent());
    }
    
    void test_get_cell()
    {
        xlnt::worksheet ws(wb_);
        auto cell = ws.get_cell(xlnt::cell_reference(1, 1));
        TS_ASSERT_EQUALS(cell.get_reference(), "A1");
    }
    
    void test_worksheet_dimension()
    {
        xlnt::worksheet ws(wb_);
        
        TS_ASSERT_EQUALS("A1:A1", ws.calculate_dimension());
        ws.get_cell("B12").set_value("AAA");
        TS_ASSERT_EQUALS("B12:B12", ws.calculate_dimension());
    }
    
    void test_squared_range()
    {
        xlnt::worksheet ws(wb_);
        
        const std::vector<std::vector<xlnt::string>> expected =
        {
            { "A1", "B1", "C1" },
            { "A2", "B2", "C2" },
            { "A3", "B3", "C3" },
            { "A4", "B4", "C4" }
        };
        
        auto rows = ws.get_squared_range(1, 1, 3, 4);
        auto expected_row_iter = expected.begin();
        
        for(auto row : rows)
        {
            auto expected_cell_iter = (*expected_row_iter).begin();
            
            for(auto cell : row)
            {
                TS_ASSERT_EQUALS(cell.get_reference(), *expected_cell_iter);
                expected_cell_iter++;
            }
            
            expected_row_iter++;
        }
    }
    
    void test_iter_rows_1()
    {
        std::size_t row = 0;
        std::size_t column = 0;
        xlnt::string coordinate = "A1";
        
        xlnt::worksheet ws(wb_);
        
        ws.get_cell("A1").set_value("first");
        ws.get_cell("C9").set_value("last");
        
        TS_ASSERT_EQUALS(ws.calculate_dimension(), "A1:C9");
        TS_ASSERT_EQUALS(ws.rows()[row][column].get_reference(), coordinate);
        
        row = 8;
        column = 2;
        coordinate = "C9";
        
        TS_ASSERT_EQUALS(ws.rows()[row][column].get_reference(), coordinate);
    }
    
    void test_iter_rows_2()
    {
        xlnt::worksheet ws(wb_);
        
        const std::vector<std::vector<xlnt::string>> expected =
        {
            { "A1", "B1", "C1" },
            { "A2", "B2", "C2" },
            { "A3", "B3", "C3" },
            { "A4", "B4", "C4" }
        };
    
        auto rows = ws.rows("A1:C4");
        auto expected_row_iter = expected.begin();
        
        for(auto row : rows)
        {
            auto expected_cell_iter = (*expected_row_iter).begin();
            
            for(auto cell : row)
            {
                TS_ASSERT_EQUALS(cell.get_reference(), *expected_cell_iter);
                expected_cell_iter++;
            }
            
            expected_row_iter++;
        }
    }
    
    void test_iter_rows_offset()
    {
        xlnt::worksheet ws(wb_);
        auto rows = ws.rows("A1:C4", 1, 3);
        
        const std::vector<std::vector<xlnt::string>> expected =
        {
            { "D2", "E2", "F2" },
            { "D3", "E3", "F3" },
            { "D4", "E4", "F4" },
            { "D5", "E5", "F5" }
        };
    
        auto expected_row_iter = expected.begin();
        
        for(auto row : rows)
        {
            auto expected_cell_iter = (*expected_row_iter).begin();
            
            for(auto cell : row)
            {
                TS_ASSERT_EQUALS(cell.get_reference(), *expected_cell_iter);
                expected_cell_iter++;
            }
            
            expected_row_iter++;
        }
    }
    
    void test_get_named_range()
    {
        xlnt::worksheet ws(wb_);
        wb_.create_named_range("test_range", ws, "C5");
        auto xlrange = ws.get_named_range("test_range");
        TS_ASSERT_EQUALS(1, xlrange.length());
        TS_ASSERT_EQUALS(1, xlrange[0].num_cells());
        TS_ASSERT_EQUALS(5, xlrange[0][0].get_row());
    }
    
    void test_get_bad_named_range()
    {
        xlnt::worksheet ws(wb_);
        TS_ASSERT_THROWS(ws.get_named_range("bad_range"), xlnt::named_range_exception);
    }
    
    void test_get_named_range_wrong_sheet()
    {
        xlnt::worksheet ws1(wb_);
        xlnt::worksheet ws2(wb_);
        wb_.create_named_range("wrong_sheet_range", ws1, "C5");
        TS_ASSERT_THROWS(ws2.get_named_range("wrong_sheet_range"), xlnt::named_range_exception);
    }
    
    void test_cell_alternate_coordinates()
    {
        xlnt::worksheet ws(wb_);
        auto cell = ws.get_cell(xlnt::cell_reference(4, 8));
        TS_ASSERT_EQUALS("D8", cell.get_reference().to_string());
    }
    
    void test_cell_range_name()
    {
        xlnt::worksheet ws(wb_);
        wb_.create_named_range("test_range_single", ws, "B12");
        auto c_range_name = ws.get_named_range("test_range_single");
        auto c_range_coord = ws.get_range("B12");
        auto c_cell = ws.get_cell("B12");
        TS_ASSERT_EQUALS(c_range_coord, c_range_name);
        TS_ASSERT(c_range_coord[0][0] == c_cell);
    }
    
    void test_garbage_collect()
    {
        xlnt::worksheet ws(wb_);
        
        ws.get_cell("A1");
        ws.get_cell("B2").set_value("0");
        ws.get_cell("C4").set_value(0);
        xlnt::comment(ws.get_cell("D1"), "Comment", "Comment");
        
        ws.garbage_collect();
        
        auto cell_collection = ws.get_cell_collection();
        std::set<xlnt::cell> cells(cell_collection.begin(), cell_collection.end());
        std::set<xlnt::cell> expected = {ws.get_cell("B2"), ws.get_cell("C4"), ws.get_cell("D1")};
        
        // Set difference
        std::set<xlnt::cell> difference;
        
        for(auto a : expected)
        {
            if(cells.find(a) == cells.end())
            {
                difference.insert(a);
            }
        }
        
        for(auto a : cells)
        {
            if(expected.find(a) == expected.end())
            {
                difference.insert(a);
            }
        }
        
        TS_ASSERT(difference.empty());
    }
    
    void test_hyperlink_value()
    {
        xlnt::worksheet ws(wb_);
        ws.get_cell("A1").set_hyperlink("http://test.com");
        TS_ASSERT_EQUALS("http://test.com", ws.get_cell("A1").get_value<xlnt::string>());
        ws.get_cell("A1").set_value("test");
        TS_ASSERT_EQUALS("test", ws.get_cell("A1").get_value<xlnt::string>());
        TS_ASSERT_EQUALS(ws.get_cell("A1").get_hyperlink().get_target_uri(), "http://test.com");
    }
    
    void test_hyperlink_relationships()
    {
        xlnt::worksheet ws(wb_);
        TS_ASSERT_EQUALS(ws.get_relationships().size(), 0);
        
        ws.get_cell("A1").set_hyperlink("http://test.com");
        TS_ASSERT_EQUALS(ws.get_relationships().size(), 1);
        
        ws.get_cell("A2").set_hyperlink("http://test2.com");
        TS_ASSERT_EQUALS(ws.get_relationships().size(), 2);
    }
    
    void test_append()
    {
        xlnt::worksheet ws(wb_);
        ws.append(std::vector<xlnt::string> {"value"});
        TS_ASSERT_EQUALS("value", ws.get_cell("A1").get_value<xlnt::string>());
    }
    
    void test_append_list()
    {
        xlnt::worksheet ws(wb_);
        
        ws.append(std::vector<xlnt::string> {"This is A1", "This is B1"});
        
        TS_ASSERT_EQUALS("This is A1", ws.get_cell("A1").get_value<xlnt::string>());
        TS_ASSERT_EQUALS("This is B1", ws.get_cell("B1").get_value<xlnt::string>());
    }
    
    void test_append_dict_letter()
    {
        xlnt::worksheet ws(wb_);
        
        const std::unordered_map<xlnt::string, xlnt::string> dict_letter =
        {
            { "A", "This is A1" },
            { "C", "This is C1" }
        };
        
        ws.append(dict_letter);
        
        TS_ASSERT_EQUALS("This is A1", ws.get_cell("A1").get_value<xlnt::string>());
        TS_ASSERT_EQUALS("This is C1", ws.get_cell("C1").get_value<xlnt::string>());
    }
    
    void test_append_dict_index()
    {
        xlnt::worksheet ws(wb_);
        
        const std::unordered_map<int, xlnt::string> dict_index =
        {
            { 1, "This is A1" },
            { 3, "This is C1" }
        };
        
        ws.append(dict_index);
        
        TS_ASSERT_EQUALS("This is A1", ws.get_cell("A1").get_value<xlnt::string>());
        TS_ASSERT_EQUALS("This is C1", ws.get_cell("C1").get_value<xlnt::string>());
    }
    
    void _test_bad_append()
    {
    }
    
    void _test_append_range()
    {

    }
    
    void test_append_iterator()
    {
        std::vector<int> range;
        
        for(int i = 0; i < 30; i++)
        {
            range.push_back(i);
        }
        
        xlnt::worksheet ws(wb_);
        ws.append(range.begin(), range.end());
        
        TS_ASSERT_EQUALS(ws[xlnt::cell_reference("AD1")].get_value<int>(), 29);
    }
    
    void test_append_2d_list()
    {
        xlnt::worksheet ws(wb_);
        
        ws.append(std::vector<xlnt::string> {"This is A1", "This is B1"});
        ws.append(std::vector<xlnt::string> {"This is A2", "This is B2"});
        
        auto vals = ws.get_range("A1:B2");
        
        TS_ASSERT_EQUALS(vals[0][0].get_value<xlnt::string>(), "This is A1");
        TS_ASSERT_EQUALS(vals[0][1].get_value<xlnt::string>(), "This is B1");
        TS_ASSERT_EQUALS(vals[1][0].get_value<xlnt::string>(), "This is A2");
        TS_ASSERT_EQUALS(vals[1][1].get_value<xlnt::string>(), "This is B2");
    }
    
    void _test_append_cell()
    {
        // Right now, a cell cannot be created without a parent worksheet.
        // This should be possible by instantiating a cell_impl, but the
        // memory management is complicated.
        /*
        xlnt::worksheet ws(wb_);
        xlnt::cell cell;
        cell.set_value(25);
        
        ws.append({ cell });
        
        TS_ASSERT_EQUALS(cell.get_value<int>(), 25);
         */
    }
    
    void test_rows()
    {
        xlnt::worksheet ws(wb_);
        
        ws.get_cell("A1").set_value("first");
        ws.get_cell("C9").set_value("last");
        
        auto rows = ws.rows();
        
        TS_ASSERT_EQUALS(rows.length(), 9);
        
        auto first_row = rows[0];
        auto last_row = rows[8];
        
        TS_ASSERT_EQUALS(first_row[0].get_value<xlnt::string>(), "first");
        TS_ASSERT_EQUALS(first_row[0].get_reference(), "A1");
        TS_ASSERT_EQUALS(last_row[2].get_value<xlnt::string>(), "last");
    }
    
    void test_no_cols()
    {
        xlnt::worksheet ws(wb_);
        
        TS_ASSERT_EQUALS(ws.columns().length(), 1);
        TS_ASSERT_EQUALS(ws.columns()[0].length(), 1);
    }
    
    void test_cols()
    {
        xlnt::worksheet ws(wb_);
        
        ws.get_cell("A1").set_value("first");
        ws.get_cell("C9").set_value("last");
        
        auto cols = ws.columns();
        
        TS_ASSERT_EQUALS(cols.length(), 3);
        
        TS_ASSERT_EQUALS(cols[0][0].get_value<xlnt::string>(), "first");
        TS_ASSERT_EQUALS(cols[2][8].get_value<xlnt::string>(), "last");
    }
    
    void test_auto_filter()
    {
        xlnt::worksheet ws(wb_);
        
        ws.auto_filter(ws.get_range("a1:f1"));
        TS_ASSERT_EQUALS(ws.get_auto_filter(), "A1:F1");
        
        ws.unset_auto_filter();
        TS_ASSERT_EQUALS(ws.has_auto_filter(), false);
        
        ws.auto_filter("c1:g9");
        TS_ASSERT_EQUALS(ws.get_auto_filter(), "C1:G9");
    }
    
    void test_getitem()
    {
        xlnt::worksheet ws(wb_);
        xlnt::cell cell = ws[xlnt::cell_reference("A1")];
        TS_ASSERT_EQUALS(cell.get_reference().to_string(), "A1");
        TS_ASSERT_EQUALS(cell.get_data_type(), xlnt::cell::type::null);
    }
    
    void test_setitem()
    {
        xlnt::worksheet ws(wb_);
        ws[xlnt::cell_reference("A1")].set_value(5);
        TS_ASSERT(ws[xlnt::cell_reference("A1")].get_value<int>() == 5);
    }
    
    void test_getslice()
    {
        xlnt::worksheet ws(wb_);
        auto cell_range = ws("A1", "B2");
        TS_ASSERT_EQUALS(cell_range[0][0], ws.get_cell("A1"));
        TS_ASSERT_EQUALS(cell_range[1][0], ws.get_cell("A2"));
        TS_ASSERT_EQUALS(cell_range[0][1], ws.get_cell("B1"));
        TS_ASSERT_EQUALS(cell_range[1][1], ws.get_cell("B2"));
    }
    
    void test_freeze()
    {
        xlnt::worksheet ws(wb_);
        
        ws.freeze_panes(ws.get_cell("b2"));
        TS_ASSERT_EQUALS(ws.get_frozen_panes(), "B2");
        
        ws.unfreeze_panes();
        TS_ASSERT(!ws.has_frozen_panes());
        
        ws.freeze_panes("c5");
        TS_ASSERT_EQUALS(ws.get_frozen_panes(), "C5");
        
        ws.freeze_panes(ws.get_cell("A1"));
        TS_ASSERT(!ws.has_frozen_panes());
    }
    
    void test_merged_cells_lookup()
    {
        xlnt::worksheet ws(wb_);
        ws.merge_cells("A1:N50");
        auto all_merged = ws.get_merged_ranges();
        TS_ASSERT_EQUALS(all_merged.size(), 1);
        auto merged = ws.get_range(all_merged[0]);
        TS_ASSERT(merged.contains("A1"));
        TS_ASSERT(merged.contains("N50"));
        TS_ASSERT(!merged.contains("A51"));
        TS_ASSERT(!merged.contains("O1"));
    }
    
    
    void test_merged_cell_ranges()
    {
        xlnt::worksheet ws(wb_);
        TS_ASSERT_EQUALS(ws.get_merged_ranges().size(), 0);
    }
    
    void test_merge_range_string()
    {
        xlnt::worksheet ws(wb_);
        ws.get_cell("A1").set_value(1);
        ws.get_cell("D4").set_value(16);
        ws.merge_cells("A1:D4");
        std::vector<xlnt::range_reference> expected = { xlnt::range_reference("A1:D4") };
        TS_ASSERT_EQUALS(ws.get_merged_ranges(), expected);
        TS_ASSERT(!ws.get_cell("D4").has_value());
    }
    
    void test_merge_coordinate()
    {
        xlnt::worksheet ws(wb_);
        ws.merge_cells(1, 1, 4, 4);
        std::vector<xlnt::range_reference> expected = { xlnt::range_reference("A1:D4") };
        TS_ASSERT_EQUALS(ws.get_merged_ranges(), expected);
    }
    
    void test_unmerge_range_string()
    {
        xlnt::worksheet ws(wb_);
        ws.merge_cells("A1:D4");
        TS_ASSERT_EQUALS(ws.get_merged_ranges().size(), 1);
        ws.unmerge_cells("A1:D4");
        TS_ASSERT_EQUALS(ws.get_merged_ranges().size(), 0);
    }
    
    void test_unmerge_coordinate()
    {
        xlnt::worksheet ws(wb_);
        ws.merge_cells("A1:D4");
        TS_ASSERT_EQUALS(ws.get_merged_ranges().size(), 1);
        ws.unmerge_cells(1, 1, 4, 4);
        TS_ASSERT_EQUALS(ws.get_merged_ranges().size(), 0);
    }
    
    void test_print_titles()
    {
        xlnt::worksheet ws(wb_);
    }
    
    void test_positioning_point()
    {
        xlnt::worksheet ws(wb_);
        TS_ASSERT_EQUALS(ws.get_point_pos(150, 40), "C3");
    }
    
    void test_positioning_roundtrip()
    {
        xlnt::worksheet ws(wb_);
        TS_ASSERT_EQUALS(ws.get_point_pos(ws.get_cell("A1").get_anchor()), xlnt::cell_reference("A1"));
        TS_ASSERT_EQUALS(ws.get_point_pos(ws.get_cell("D52").get_anchor()), xlnt::cell_reference("D52"));
        TS_ASSERT_EQUALS(ws.get_point_pos(ws.get_cell("X11").get_anchor()), xlnt::cell_reference("X11"));
    }
    
    void test_freeze_panes_horiz()
    {
        xlnt::worksheet ws(wb_);
        ws.freeze_panes("A4");
    }
    
    void test_freeze_panes_vert()
    {
        xlnt::worksheet ws(wb_);
        ws.freeze_panes("D1");
    }
    
    void test_freeze_panes_both()
    {
        xlnt::worksheet ws(wb_);
        ws.freeze_panes("D4");
    }
    
    void test_min_column()
    {
        xlnt::worksheet ws(wb_);
        TS_ASSERT_EQUALS(ws.get_lowest_column(), 1);
    }
    
    void test_max_column()
    {
        xlnt::worksheet ws(wb_);
        ws[xlnt::cell_reference("F1")].set_value(10);
        ws[xlnt::cell_reference("F2")].set_value(32);
        ws[xlnt::cell_reference("F3")].set_formula("=F1+F2");
        ws[xlnt::cell_reference("A4")].set_formula("=A1+A2+A3");
        TS_ASSERT_EQUALS(ws.get_highest_column(), 6);
    }

    void test_min_row()
    {
        xlnt::worksheet ws(wb_);
        TS_ASSERT_EQUALS(ws.get_lowest_row(), 1);
    }
    
    void test_max_row()
    {
        xlnt::worksheet ws(wb_);
        ws.append();
        ws.append(std::vector<int> { 5 });
        ws.append();
        ws.append(std::vector<int> { 4 });
        TS_ASSERT_EQUALS(ws.get_highest_row(), 4);
    }
    
    // end 2.4 synchronized tests
    
    void test_new_sheet_name()
    {
        xlnt::worksheet ws = wb_.create_sheet("TestName");
        TS_ASSERT_EQUALS(ws.to_string(), "<Worksheet \"TestName\">");
    }

    void test_set_bad_title()
    {
        xlnt::string title(std::string(50, 'X').c_str());
        TS_ASSERT_THROWS(wb_.create_sheet(title), xlnt::sheet_title_exception);
    }
    
    void test_increment_title()
    {
    	auto ws1 = wb_.create_sheet("Test");
    	TS_ASSERT_EQUALS(ws1.get_title(), "Test");
    	auto ws2 = wb_.create_sheet("Test");
    	TS_ASSERT_EQUALS(ws2.get_title(), "Test1");
    }

    void test_set_bad_title_character()
    {
        TS_ASSERT_THROWS(wb_.create_sheet("["), xlnt::sheet_title_exception);
        TS_ASSERT_THROWS(wb_.create_sheet("]"), xlnt::sheet_title_exception);
        TS_ASSERT_THROWS(wb_.create_sheet("*"), xlnt::sheet_title_exception);
        TS_ASSERT_THROWS(wb_.create_sheet(":"), xlnt::sheet_title_exception);
        TS_ASSERT_THROWS(wb_.create_sheet("?"), xlnt::sheet_title_exception);
        TS_ASSERT_THROWS(wb_.create_sheet("/"), xlnt::sheet_title_exception);
        TS_ASSERT_THROWS(wb_.create_sheet("\\"), xlnt::sheet_title_exception);
    }
    
    void test_unique_sheet_title()
    {
        auto ws = wb_.create_sheet("AGE");
        TS_ASSERT_EQUALS(ws.unique_sheet_name("GE"), "GE");
    }


    void test_worksheet_range()
    {
        xlnt::worksheet ws(wb_);
        auto xlrange = ws.get_range("A1:C4");
        TS_ASSERT_EQUALS(4, xlrange.length());
        TS_ASSERT_EQUALS(3, xlrange[0].num_cells());
    }

    void test_cell_offset()
    {
        xlnt::worksheet ws(wb_);
        TS_ASSERT_EQUALS("C17", ws.get_cell(xlnt::cell_reference("B15").make_offset(1, 2)).get_reference().to_string());
    }

    void test_range_offset()
    {
        xlnt::worksheet ws(wb_);
        auto xlrange = ws.get_range(xlnt::range_reference("A1:C4").make_offset(3, 1));
        TS_ASSERT_EQUALS(4, xlrange.length());
        TS_ASSERT_EQUALS(3, xlrange[0].num_cells());
        TS_ASSERT_EQUALS("D2", xlrange[0][0].get_reference().to_string());
    }

    void test_bad_relationship_type()
    {
        xlnt::relationship rel("bad");
    }

    void test_write_empty()
    {
        xlnt::worksheet ws(wb_);
        
        xlnt::worksheet_serializer serializer(ws);
        auto observed = serializer.write_worksheet();

        auto expected_string =
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "  <sheetPr>"
        "    <outlinePr summaryRight=\"1\" summaryBelow=\"1\"/>"
        "  </sheetPr>"
        "  <dimension ref=\"A1:A1\"/>"
        "  <sheetViews>"
        "    <sheetView workbookViewId=\"0\">"
        "      <selection sqref=\"A1\" activeCell=\"A1\"/>"
        "    </sheetView>"
        "  </sheetViews>"
        "  <sheetFormatPr baseColWidth=\"10\" defaultRowHeight=\"15\"/>"
        "  <sheetData/>"
        "  <pageMargins left=\"0.75\" right=\"0.75\" top=\"1\" bottom=\"1\" header=\"0.5\" footer=\"0.5\"/>"
        "</worksheet>";

        xlnt::xml_document expected;
        expected.from_string(expected_string);

        TS_ASSERT(Helper::compare_xml(expected, observed));
    }
    
    void test_page_margins()
    {
        xlnt::worksheet ws(wb_);

        ws.get_page_margins().set_left(2);
        ws.get_page_margins().set_right(2);
        ws.get_page_margins().set_top(2);
        ws.get_page_margins().set_bottom(2);
        ws.get_page_margins().set_header(1.5);
        ws.get_page_margins().set_footer(1.5);
        
        xlnt::worksheet_serializer serializer(ws);
        auto observed = serializer.write_worksheet();

        auto expected_string = 
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "  <sheetPr>"
        "    <outlinePr summaryRight=\"1\" summaryBelow=\"1\"/>"
        "  </sheetPr>"
        "  <dimension ref=\"A1:A1\"/>"
        "  <sheetViews>"
        "    <sheetView workbookViewId=\"0\">"
        "      <selection sqref=\"A1\" activeCell=\"A1\"/>"
        "    </sheetView>"
        "  </sheetViews>"
        "  <sheetFormatPr baseColWidth=\"10\" defaultRowHeight=\"15\"/>"
        "  <sheetData/>"
        "  <pageMargins left=\"2\" right=\"2\" top=\"2\" bottom=\"2\" header=\"1.5\" footer=\"1.5\"/>"
        "</worksheet>";

        xlnt::xml_document expected;
        expected.from_string(expected_string);
        
        TS_ASSERT(Helper::compare_xml(expected, observed));
    }
   
    void test_merge()
    {
        xlnt::workbook clean_wb; // for shared strings
        xlnt::worksheet ws(clean_wb);
        
        auto expected_string1 = 
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "  <sheetPr>"
        "    <outlinePr summaryRight=\"1\" summaryBelow=\"1\"/>"
        "  </sheetPr>"
        "  <dimension ref=\"A1:B1\"/>"
        "  <sheetViews>"
        "    <sheetView workbookViewId=\"0\">"
        "      <selection sqref=\"A1\" activeCell=\"A1\"/>"
        "    </sheetView>"
        "  </sheetViews>"
        "  <sheetFormatPr baseColWidth=\"10\" defaultRowHeight=\"15\"/>"
        "  <sheetData>"
        "    <row r=\"1\" spans=\"1:2\">"
        "      <c r=\"A1\" t=\"s\">"
        "        <v>0</v>"
        "      </c>"
        "      <c r=\"B1\" t=\"s\">"
        "        <v>1</v>"
        "      </c>"
        "    </row>"
        "  </sheetData>"
        "  <pageMargins left=\"0.75\" right=\"0.75\" top=\"1\" bottom=\"1\" header=\"0.5\" footer=\"0.5\"/>"
        "</worksheet>";

        ws.get_cell("A1").set_value("Cell A1");
        ws.get_cell("B1").set_value("Cell B1");

        {
            xlnt::worksheet_serializer serializer(ws);
            auto observed = serializer.write_worksheet();

            xlnt::xml_document expected;
            expected.from_string(expected_string1);
            
            TS_ASSERT(Helper::compare_xml(expected, observed));
        }

        ws.merge_cells("A1:B1");

        auto expected_string2 = 
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "  <sheetPr>"
        "    <outlinePr summaryRight=\"1\" summaryBelow=\"1\"/>"
        "  </sheetPr>"
        "  <dimension ref=\"A1:B1\"/>"
        "  <sheetViews>"
        "    <sheetView workbookViewId=\"0\">"
        "      <selection sqref=\"A1\" activeCell=\"A1\"/>"
        "    </sheetView>"
        "  </sheetViews>"
        "  <sheetFormatPr baseColWidth=\"10\" defaultRowHeight=\"15\"/>"
        "  <sheetData>"
        "    <row r=\"1\" spans=\"1:2\">"
        "      <c r=\"A1\" t=\"s\">"
        "        <v>0</v>"
        "      </c>"
        "      <c r=\"B1\" t=\"s\"/>"
        "    </row>"
        "  </sheetData>"
        " <mergeCells count=\"1\">"
        "    <mergeCell ref=\"A1:B1\"/>"
        "  </mergeCells>"
        "  <pageMargins left=\"0.75\" right=\"0.75\" top=\"1\" bottom=\"1\" header=\"0.5\" footer=\"0.5\"/>"
        "</worksheet>";

        {
            xlnt::worksheet_serializer serializer(ws);
            auto observed = serializer.write_worksheet();

            xlnt::xml_document expected;
            expected.from_string(expected_string2);
            
            TS_ASSERT(Helper::compare_xml(expected, observed));
        }

        ws.unmerge_cells("A1:B1");

        auto expected_string3 = 
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "  <sheetPr>"
        "    <outlinePr summaryRight=\"1\" summaryBelow=\"1\"/>"
        "  </sheetPr>"
        "  <dimension ref=\"A1:B1\"/>"
        "  <sheetViews>"
        "    <sheetView workbookViewId=\"0\">"
        "      <selection sqref=\"A1\" activeCell=\"A1\"/>"
        "    </sheetView>"
        "  </sheetViews>"
        "  <sheetFormatPr baseColWidth=\"10\" defaultRowHeight=\"15\"/>"
        "  <sheetData>"
        "    <row r=\"1\" spans=\"1:2\">"
        "      <c r=\"A1\" t=\"s\">"
        "        <v>0</v>"
        "      </c>"
        "      <c r=\"B1\" t=\"s\"/>"
        "    </row>"
        "  </sheetData>"
        "  <pageMargins left=\"0.75\" right=\"0.75\" top=\"1\" bottom=\"1\" header=\"0.5\" footer=\"0.5\"/>"
        "</worksheet>";

        {
            xlnt::worksheet_serializer serializer(ws);
            auto observed = serializer.write_worksheet();
            
            xlnt::xml_document expected;
            expected.from_string(expected_string3);
            
            TS_ASSERT(Helper::compare_xml(expected, observed));
        }
    }

    void test_printer_settings()
    {
        xlnt::worksheet ws(wb_);

        ws.get_page_setup().set_orientation(xlnt::page_setup::orientation::landscape);
        ws.get_page_setup().set_paper_size(xlnt::page_setup::paper_size::tabloid);
        ws.get_page_setup().set_fit_to_page(true);
        ws.get_page_setup().set_fit_to_height(false);
        ws.get_page_setup().set_fit_to_width(true);
        ws.get_page_setup().set_horizontal_centered(true);
        ws.get_page_setup().set_vertical_centered(true);

        xlnt::worksheet_serializer serializer(ws);
        auto observed = serializer.write_worksheet();

        auto expected_string =
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "   <sheetPr>"
        "      <outlinePr summaryRight=\"1\" summaryBelow=\"1\" />"
        "      <pageSetUpPr fitToPage=\"1\" />"
        "   </sheetPr>"
        "   <dimension ref=\"A1:A1\" />"
        "   <sheetViews>"
        "      <sheetView workbookViewId=\"0\">"
        "         <selection sqref=\"A1\" activeCell=\"A1\" />"
        "      </sheetView>"
        "   </sheetViews>"
        "   <sheetFormatPr baseColWidth=\"10\" defaultRowHeight=\"15\" />"
        "   <sheetData />"
        "   <printOptions horizontalCentered=\"1\" verticalCentered=\"1\" />"
        "   <pageMargins left=\"0.75\" right=\"0.75\" top=\"1\" bottom=\"1\" header=\"0.5\" footer=\"0.5\" />"
        "   <pageSetup orientation=\"landscape\" paperSize=\"3\" fitToHeight=\"0\" fitToWidth=\"1\" />"
        "</worksheet>";
        
        xlnt::xml_document expected;
        expected.from_string(expected_string);
        
        TS_ASSERT(Helper::compare_xml(expected, observed));
    }
    
    void test_header_footer()
    {
    	auto ws = wb_.create_sheet();
    	ws.get_header_footer().get_left_header().set_text("Left Header Text");
        ws.get_header_footer().get_center_header().set_text("Center Header Text");
        ws.get_header_footer().get_center_header().set_font_name("Arial,Regular");
        ws.get_header_footer().get_center_header().set_font_size(6);
        ws.get_header_footer().get_center_header().set_font_color("445566");
        ws.get_header_footer().get_right_header().set_text("Right Header Text");
        ws.get_header_footer().get_right_header().set_font_name("Arial,Bold");
        ws.get_header_footer().get_right_header().set_font_size(8);
        ws.get_header_footer().get_right_header().set_font_color("112233");
        ws.get_header_footer().get_left_footer().set_text("Left Footer Text\nAnd &[Date] and &[Time]");
        ws.get_header_footer().get_left_footer().set_font_name("Times New Roman,Regular");
        ws.get_header_footer().get_left_footer().set_font_size(10);
        ws.get_header_footer().get_left_footer().set_font_color("445566");
        ws.get_header_footer().get_center_footer().set_text("Center Footer Text &[Path]&[File] on &[Tab]");
        ws.get_header_footer().get_center_footer().set_font_name("Times New Roman,Bold");
        ws.get_header_footer().get_center_footer().set_font_size(12);
        ws.get_header_footer().get_center_footer().set_font_color("778899");
        ws.get_header_footer().get_right_footer().set_text("Right Footer Text &[Page] of &[Pages]");
        ws.get_header_footer().get_right_footer().set_font_name("Times New Roman,Italic");
        ws.get_header_footer().get_right_footer().set_font_size(14);
        ws.get_header_footer().get_right_footer().set_font_color("AABBCC");
        
        xlnt::string expected_xml_string =
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "  <sheetPr>"
        "    <outlinePr summaryRight=\"1\" summaryBelow=\"1\"/>"
        "  </sheetPr>"
        "  <dimension ref=\"A1:A1\"/>"
        "  <sheetViews>"
        "    <sheetView workbookViewId=\"0\">"
        "      <selection sqref=\"A1\" activeCell=\"A1\"/>"
        "    </sheetView>"
        "  </sheetViews>"
        "  <sheetFormatPr baseColWidth=\"10\" defaultRowHeight=\"15\"/>"
        "  <sheetData/>"
        "  <pageMargins left=\"0.75\" right=\"0.75\" top=\"1\" bottom=\"1\" header=\"0.5\" footer=\"0.5\"/>"
        "  <headerFooter>"
        "    <oddHeader>&amp;L&amp;\"Calibri,Regular\"&amp;K000000Left Header Text&amp;C&amp;\"Arial,Regular\"&amp;6&amp;K445566Center Header Text&amp;R&amp;\"Arial,Bold\"&amp;8&amp;K112233Right Header Text</oddHeader>"
        "    <oddFooter>&amp;L&amp;\"Times New Roman,Regular\"&amp;10&amp;K445566Left Footer Text_x000D_And &amp;D and &amp;T&amp;C&amp;\"Times New Roman,Bold\"&amp;12&amp;K778899Center Footer Text &amp;Z&amp;F on &amp;A&amp;R&amp;\"Times New Roman,Italic\"&amp;14&amp;KAABBCCRight Footer Text &amp;P of &amp;N</oddFooter>"
        "  </headerFooter>"
        "</worksheet>";
        
        {
            xlnt::worksheet_serializer serializer(ws);
            auto observed = serializer.write_worksheet();
            
            xlnt::xml_document expected;
            expected.from_string(expected_xml_string);
            
            TS_ASSERT(Helper::compare_xml(expected, observed));
        }
        
        ws = wb_.create_sheet();
        
        expected_xml_string =
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "  <sheetPr>"
        "    <outlinePr summaryRight=\"1\" summaryBelow=\"1\"/>"
        "  </sheetPr>"
        "  <dimension ref=\"A1:A1\"/>"
        "  <sheetViews>"
        "    <sheetView workbookViewId=\"0\">"
        "      <selection sqref=\"A1\" activeCell=\"A1\"/>"
        "    </sheetView>"
        "  </sheetViews>"
        "  <sheetFormatPr baseColWidth=\"10\" defaultRowHeight=\"15\"/>"
        "  <sheetData/>"
        "  <pageMargins left=\"0.75\" right=\"0.75\" top=\"1\" bottom=\"1\" header=\"0.5\" footer=\"0.5\"/>"
        "</worksheet>";
        
        {
            xlnt::worksheet_serializer serializer(ws);
            auto observed = serializer.write_worksheet();
            
            xlnt::xml_document expected;
            expected.from_string(expected_xml_string);
            
            TS_ASSERT(Helper::compare_xml(expected, observed));
        }
    }
    
    void test_page_setup()
    {
    	xlnt::page_setup p;
    	TS_ASSERT_EQUALS(p.get_scale(), 1);
    	p.set_scale(2);
    	TS_ASSERT_EQUALS(p.get_scale(), 2);
    }
    
    void test_page_options()
    {
    	xlnt::page_setup p;
    	TS_ASSERT(!p.get_horizontal_centered());
        TS_ASSERT(!p.get_vertical_centered());
    	p.set_horizontal_centered(true);
    	p.set_vertical_centered(true);
        TS_ASSERT(p.get_horizontal_centered());
        TS_ASSERT(p.get_vertical_centered());
    }

private:
    xlnt::workbook wb_;
};
