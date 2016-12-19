#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/worksheet/header_footer.hpp>
#include <xlnt/worksheet/worksheet.hpp>
#include <xlnt/workbook/workbook.hpp>

class test_worksheet : public CxxTest::TestSuite
{
public:
    void test_new_worksheet()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        TS_ASSERT(ws.workbook() == wb);
    }
    
    void test_cell()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        auto cell = ws.cell(xlnt::cell_reference(1, 1));
        TS_ASSERT_EQUALS(cell.reference(), "A1");
    }
    
    void test_invalid_cell()
    {
        TS_ASSERT_THROWS(xlnt::cell_reference(xlnt::column_t((xlnt::column_t::index_t)0), 0),
            xlnt::invalid_cell_reference);
    }
    
    void test_worksheet_dimension()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        
        TS_ASSERT_EQUALS("A1:A1", ws.calculate_dimension());
        ws.cell("B12").value("AAA");
        TS_ASSERT_EQUALS("B12:B12", ws.calculate_dimension());
    }
    
    void test_fill_rows()
    {
        std::size_t row = 0;
        std::size_t column = 0;
        std::string coordinate = "A1";
        
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        
        ws.cell("A1").value("first");
        ws.cell("C9").value("last");
        
        TS_ASSERT_EQUALS(ws.calculate_dimension(), "A1:C9");
        TS_ASSERT_EQUALS(ws.rows()[row][column].reference(), coordinate);
        
        row = 8;
        column = 2;
        coordinate = "C9";
        
        TS_ASSERT_EQUALS(ws.rows()[row][column].reference(), coordinate);
    }
    
    void test_iter_rows()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        
        const std::vector<std::vector<std::string>> expected =
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
                TS_ASSERT_EQUALS(cell.reference(), *expected_cell_iter);
                expected_cell_iter++;
            }
            
            expected_row_iter++;
        }
    }
    
    void test_iter_rows_offset()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        auto rows = ws.rows("A1:C4", 1, 3);
        
        const std::vector<std::vector<std::string>> expected =
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
                TS_ASSERT_EQUALS(cell.reference(), *expected_cell_iter);
                expected_cell_iter++;
            }
            
            expected_row_iter++;
        }
    }

    void test_iter_rows_offset_int_int()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        auto rows = ws.rows(1, 3);

        const std::vector<std::vector<std::string>> expected =
        {
            { "D2", "E2", "F2" },
            { "D3", "E3", "F3" },
            { "D4", "E4", "F4" },
            { "D5", "E5", "F5" }
        };

        auto expected_row_iter = expected.begin();

        for (auto row : rows)
        {
            auto expected_cell_iter = (*expected_row_iter).begin();

            for (auto cell : row)
            {
                TS_ASSERT_EQUALS(cell.reference(), *expected_cell_iter);
                expected_cell_iter++;
            }

            expected_row_iter++;
        }
    }
    
    void test_get_named_range()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        wb.create_named_range("test_range", ws, "C5");
        auto xlrange = ws.named_range("test_range");
        TS_ASSERT_EQUALS(1, xlrange.length());
        TS_ASSERT_EQUALS(1, xlrange[0].length());
        TS_ASSERT_EQUALS(5, xlrange[0][0].row());

        ws.create_named_range("test_range2", "C6");
        auto xlrange2 = ws.named_range("test_range2");
        TS_ASSERT_EQUALS(1, xlrange2.length());
        TS_ASSERT_EQUALS(1, xlrange2[0].length());
        TS_ASSERT_EQUALS(6, xlrange2[0][0].row());
    }
    
    void test_get_bad_named_range()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        TS_ASSERT_THROWS(ws.named_range("bad_range"), xlnt::key_not_found);
    }
    
    void test_get_named_range_wrong_sheet()
    {
        xlnt::workbook wb;
        wb.create_sheet();
        wb.create_sheet();
        auto ws1 = wb[0];
        auto ws2 = wb[1];
        wb.create_named_range("wrong_sheet_range", ws1, "C5");
        TS_ASSERT_THROWS(ws2.named_range("wrong_sheet_range"), xlnt::key_not_found);
    }
    
    void test_remove_named_range_bad()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        TS_ASSERT_THROWS(ws.remove_named_range("bad_range"), std::runtime_error);
    }

    void test_cell_alternate_coordinates()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        auto cell = ws.cell(xlnt::cell_reference(4, 8));
        TS_ASSERT_EQUALS(cell.reference(), "D8");
    }
    
    // void test_cell_insufficient_coordinates() {}
    
    void test_cell_range_name()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        wb.create_named_range("test_range_single", ws, "B12");
        auto c_range_name = ws.named_range("test_range_single");
        auto c_range_coord = ws.range("B12");
        auto c_cell = ws.cell("B12");
        TS_ASSERT_EQUALS(c_range_coord, c_range_name);
        TS_ASSERT(c_range_coord[0][0] == c_cell);
    }
    
    void test_hyperlink_value()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        ws.cell("A1").hyperlink("http://test.com");
        TS_ASSERT_EQUALS("http://test.com", ws.cell("A1").value<std::string>());
        ws.cell("A1").value("test");
        TS_ASSERT_EQUALS("test", ws.cell("A1").value<std::string>());
        TS_ASSERT_EQUALS(ws.cell("A1").hyperlink(), "http://test.com");
    }
    
    void test_append()
    {
        xlnt::workbook wb;
		auto ws = wb.active_sheet();
        ws.append(std::vector<std::string> {"value"});
        TS_ASSERT_EQUALS("value", ws.cell("A1").value<std::string>());
    }
    
    void test_append_list()
    {
        xlnt::workbook wb;
		auto ws = wb.active_sheet();
        
        ws.append(std::vector<std::string> {"This is A1", "This is B1"});
        
        TS_ASSERT_EQUALS("This is A1", ws.cell("A1").value<std::string>());
        TS_ASSERT_EQUALS("This is B1", ws.cell("B1").value<std::string>());
    }
    
    void test_append_dict_letter()
    {
        xlnt::workbook wb;
		auto ws = wb.active_sheet();
        
        const std::unordered_map<std::string, std::string> dict_letter =
        {
            { "A", "This is A1" },
            { "C", "This is C1" }
        };
        
        ws.append(dict_letter);
        
        TS_ASSERT_EQUALS("This is A1", ws.cell("A1").value<std::string>());
        TS_ASSERT_EQUALS("This is C1", ws.cell("C1").value<std::string>());
    }
    
    void test_append_dict_index()
    {
        xlnt::workbook wb;
		auto ws = wb.active_sheet();
        
        const std::unordered_map<int, std::string> dict_index =
        {
            { 1, "This is A1" },
            { 3, "This is C1" }
        };
        
        ws.append(dict_index);
        
        TS_ASSERT_EQUALS("This is A1", ws.cell("A1").value<std::string>());
        TS_ASSERT_EQUALS("This is C1", ws.cell("C1").value<std::string>());
    }
    
    void test_append_iterator()
    {
        std::vector<int> range;
        
        for(int i = 0; i < 30; i++)
        {
            range.push_back(i);
        }
        
        xlnt::workbook wb;
		auto ws = wb.active_sheet();
        ws.append(range.begin(), range.end());
        
        TS_ASSERT_EQUALS(ws[xlnt::cell_reference("AD1")].value<int>(), 29);
    }
    
    void test_append_2d_list()
    {
        xlnt::workbook wb;
		auto ws = wb.active_sheet();
        
        ws.append(std::vector<std::string> {"This is A1", "This is B1"});
        ws.append(std::vector<std::string> {"This is A2", "This is B2"});
        
        auto vals = ws.range("A1:B2");
        
        TS_ASSERT_EQUALS(vals[0][0].value<std::string>(), "This is A1");
        TS_ASSERT_EQUALS(vals[0][1].value<std::string>(), "This is B1");
        TS_ASSERT_EQUALS(vals[1][0].value<std::string>(), "This is A2");
        TS_ASSERT_EQUALS(vals[1][1].value<std::string>(), "This is B2");
    }

    void test_rows()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        
        ws.cell("A1").value("first");
        ws.cell("C9").value("last");
        
        auto rows = ws.rows();
        
        TS_ASSERT_EQUALS(rows.length(), 9);
        
        auto first_row = rows[0];
        auto last_row = rows[8];
        
        TS_ASSERT_EQUALS(first_row[0].value<std::string>(), "first");
        TS_ASSERT_EQUALS(first_row[0].reference(), "A1");
        TS_ASSERT_EQUALS(last_row[2].value<std::string>(), "last");
    }
    
    void test_no_rows()
    {
        xlnt::workbook wb;
		auto ws = wb.active_sheet();
        
        TS_ASSERT_EQUALS(ws.rows().length(), 1);
        TS_ASSERT_EQUALS(ws.rows()[0].length(), 1);
    }
    
    void test_no_cols()
    {
        xlnt::workbook wb;
		auto ws = wb.active_sheet();
        
        TS_ASSERT_EQUALS(ws.columns().length(), 1);
        TS_ASSERT_EQUALS(ws.columns()[0].length(), 1);
    }
    
    void test_one_cell()
    {
        xlnt::workbook wb;
		auto ws = wb.active_sheet();
        
        auto cell = ws.cell("A1");
        
        TS_ASSERT_EQUALS(ws.columns().length(), 1);
        TS_ASSERT_EQUALS(ws.columns()[0].length(), 1);
        TS_ASSERT_EQUALS(ws.columns()[0][0], cell);
    }
    
    // void test_by_col() {}
    
    void test_cols()
    {
        xlnt::workbook wb;
		auto ws = wb.active_sheet();
        
        ws.cell("A1").value("first");
        ws.cell("C9").value("last");
        
        auto cols = ws.columns();
        
        TS_ASSERT_EQUALS(cols.length(), 3);
        
        TS_ASSERT_EQUALS(cols[0][0].value<std::string>(), "first");
        TS_ASSERT_EQUALS(cols[2][8].value<std::string>(), "last");
    }
    
    void test_auto_filter()
    {
        xlnt::workbook wb;
		auto ws = wb.active_sheet();
        
        ws.auto_filter(ws.range("a1:f1"));
        TS_ASSERT_EQUALS(ws.auto_filter(), "A1:F1");
        
        ws.clear_auto_filter();
        TS_ASSERT(!ws.has_auto_filter());
        
        ws.auto_filter("c1:g9");
        TS_ASSERT_EQUALS(ws.auto_filter(), "C1:G9");
    }
    
    void test_getitem()
    {
        xlnt::workbook wb;
		auto ws = wb.active_sheet();
        xlnt::cell cell = ws[xlnt::cell_reference("A1")];
        TS_ASSERT_EQUALS(cell.reference().to_string(), "A1");
        TS_ASSERT_EQUALS(cell.data_type(), xlnt::cell::type::null);
    }

    void test_setitem()
    {
        xlnt::workbook wb;
		auto ws = wb.active_sheet();
        ws[xlnt::cell_reference("A12")].value(5);
        TS_ASSERT(ws[xlnt::cell_reference("A12")].value<int>() == 5);
    }
    
    void test_getslice()
    {
        xlnt::workbook wb;
		auto ws = wb.active_sheet();
        auto cell_range = ws("A1", "B2");
        TS_ASSERT_EQUALS(cell_range[0][0], ws.cell("A1"));
        TS_ASSERT_EQUALS(cell_range[1][0], ws.cell("A2"));
        TS_ASSERT_EQUALS(cell_range[0][1], ws.cell("B1"));
        TS_ASSERT_EQUALS(cell_range[1][1], ws.cell("B2"));
    }

    void test_freeze()
    {
        xlnt::workbook wb;
		auto ws = wb.active_sheet();
        
        ws.freeze_panes(ws.cell("b2"));
        TS_ASSERT_EQUALS(ws.frozen_panes(), "B2");
        
        ws.unfreeze_panes();
        TS_ASSERT(!ws.has_frozen_panes());
        
        ws.freeze_panes("c5");
        TS_ASSERT_EQUALS(ws.frozen_panes(), "C5");
        
        ws.freeze_panes(ws.cell("A1"));
        TS_ASSERT(!ws.has_frozen_panes());
    }
    
    void test_merged_cells_lookup()
    {
        xlnt::workbook wb;
		auto ws = wb.active_sheet();
        ws.cell("A2").value("test");
        ws.merge_cells("A1:N50");
        auto all_merged = ws.merged_ranges();
        TS_ASSERT_EQUALS(all_merged.size(), 1);
        auto merged = ws.range(all_merged[0]);
        TS_ASSERT(merged.contains("A1"));
        TS_ASSERT(merged.contains("N50"));
        TS_ASSERT(!merged.contains("A51"));
        TS_ASSERT(!merged.contains("O1"));
    }
    
    
    void test_merged_cell_ranges()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        TS_ASSERT_EQUALS(ws.merged_ranges().size(), 0);
    }
    
    void test_merge_range_string()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        ws.cell("A1").value(1);
        ws.cell("D4").value(16);
        ws.merge_cells("A1:D4");
        std::vector<xlnt::range_reference> expected = { xlnt::range_reference("A1:D4") };
        TS_ASSERT_EQUALS(ws.merged_ranges(), expected);
        TS_ASSERT(!ws.cell("D4").has_value());
    }
    
    void test_merge_coordinate()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        ws.merge_cells(1, 1, 4, 4);
        std::vector<xlnt::range_reference> expected = { xlnt::range_reference("A1:D4") };
        TS_ASSERT_EQUALS(ws.merged_ranges(), expected);
    }
    
    void test_unmerge_bad()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();

        TS_ASSERT_THROWS(ws.unmerge_cells("A1:D3"), std::runtime_error);
    }

    void test_unmerge_range_string()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        ws.merge_cells("A1:D4");
        TS_ASSERT_EQUALS(ws.merged_ranges().size(), 1);
        ws.unmerge_cells("A1:D4");
        TS_ASSERT_EQUALS(ws.merged_ranges().size(), 0);
    }

    void test_unmerge_coordinate()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        ws.merge_cells("A1:D4");
        TS_ASSERT_EQUALS(ws.merged_ranges().size(), 1);
        ws.unmerge_cells(1, 1, 4, 4);
        TS_ASSERT_EQUALS(ws.merged_ranges().size(), 0);
    }

    void test_print_titles_old()
    {
        xlnt::workbook wb;

        auto ws = wb.active_sheet();
        ws.print_title_rows(3);
        TS_ASSERT_EQUALS(ws.print_titles(), "Sheet1!1:3");
        
        auto ws2 = wb.create_sheet();
        ws2.print_title_cols(4);
        TS_ASSERT_EQUALS(ws2.print_titles(), "Sheet2!A:D");
    }
    
    void test_print_titles_new()
    {
        xlnt::workbook wb;
        
        auto ws = wb.active_sheet();
        ws.print_title_rows(4);
        TS_ASSERT_EQUALS(ws.print_titles(), "Sheet1!1:4");

        auto ws2 = wb.create_sheet();
        ws2.print_title_cols("F");
        TS_ASSERT_EQUALS(ws2.print_titles(), "Sheet2!A:F");

        auto ws3 = wb.create_sheet();
        ws3.print_title_rows(2, 3);
        ws3.print_title_cols("C", "D");
        TS_ASSERT_EQUALS(ws3.print_titles(), "Sheet3!2:3,Sheet3!C:D");
    }
    
    void test_print_area()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        ws.print_area("A1:F5");
        TS_ASSERT_EQUALS(ws.print_area(), "$A$1:$F$5");
    }
    
    void test_freeze_panes_horiz()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        ws.freeze_panes("A4");
        
        auto view = ws.view();
        TS_ASSERT_EQUALS(view.selections().size(), 2);
        TS_ASSERT_EQUALS(view.selections()[0].active_cell(), "A3");
        TS_ASSERT_EQUALS(view.selections()[0].pane(), xlnt::pane_corner::bottom_left);
        TS_ASSERT_EQUALS(view.selections()[0].sqref(), "A1");
        TS_ASSERT_EQUALS(view.pane().active_pane, xlnt::pane_corner::bottom_left);
        TS_ASSERT_EQUALS(view.pane().state, xlnt::pane_state::frozen);
        TS_ASSERT_EQUALS(view.pane().top_left_cell.get(), "A4");
        TS_ASSERT_EQUALS(view.pane().y_split, 3);
    }
    
    void test_freeze_panes_vert()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        ws.freeze_panes("D1");
        
        auto view = ws.view();
        TS_ASSERT_EQUALS(view.selections().size(), 2);
        TS_ASSERT_EQUALS(view.selections()[0].active_cell(), "C1");
        TS_ASSERT_EQUALS(view.selections()[0].pane(), xlnt::pane_corner::top_right);
        TS_ASSERT_EQUALS(view.selections()[0].sqref(), "A1");
        TS_ASSERT_EQUALS(view.pane().active_pane, xlnt::pane_corner::top_right);
        TS_ASSERT_EQUALS(view.pane().state, xlnt::pane_state::frozen);
        TS_ASSERT_EQUALS(view.pane().top_left_cell.get(), "D1");
        TS_ASSERT_EQUALS(view.pane().x_split, 3);
    }
    
    void test_freeze_panes_both()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        ws.freeze_panes("D4");
        
        auto view = ws.view();
        TS_ASSERT_EQUALS(view.selections().size(), 3);
        TS_ASSERT_EQUALS(view.selections()[0].pane(), xlnt::pane_corner::top_right);
        TS_ASSERT_EQUALS(view.selections()[1].pane(), xlnt::pane_corner::bottom_left);
        TS_ASSERT_EQUALS(view.selections()[2].active_cell(), "D4");
        TS_ASSERT_EQUALS(view.selections()[2].pane(), xlnt::pane_corner::bottom_right);
        TS_ASSERT_EQUALS(view.selections()[2].sqref(), "A1");
        TS_ASSERT_EQUALS(view.pane().active_pane, xlnt::pane_corner::bottom_right);
        TS_ASSERT_EQUALS(view.pane().state, xlnt::pane_state::frozen);
        TS_ASSERT_EQUALS(view.pane().top_left_cell.get(), "D4");
        TS_ASSERT_EQUALS(view.pane().x_split, 3);
        TS_ASSERT_EQUALS(view.pane().y_split, 3);
    }
    
    void test_min_column()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        TS_ASSERT_EQUALS(ws.lowest_column(), 1);
    }
    
    void test_max_column()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        ws[xlnt::cell_reference("F1")].value(10);
        ws[xlnt::cell_reference("F2")].value(32);
        ws[xlnt::cell_reference("F3")].formula("=F1+F2");
        ws[xlnt::cell_reference("A4")].formula("=A1+A2+A3");
        TS_ASSERT_EQUALS(ws.highest_column(), 6);
    }

    void test_min_row()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        TS_ASSERT_EQUALS(ws.lowest_row(), 1);
    }
    
    void test_max_row()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        ws.append();
        ws.append(std::vector<int> { 5 });
        ws.append();
        ws.append(std::vector<int> { 4 });
        TS_ASSERT_EQUALS(ws.highest_row(), 4);
    }
    
    void test_const_iterators()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();

        ws.append({"A1", "B1", "C1"});
        ws.append({"A2", "B2", "C2"});

        const xlnt::worksheet ws_const = ws;
        const auto rows = ws_const.rows();

        const auto first_row = *rows.begin();
        const auto first_cell = *first_row.begin();
        TS_ASSERT_EQUALS(first_cell.value<std::string>(), "A1");

        const auto last_row = *(--rows.end());
        const auto last_cell = *(--last_row.end());
        TS_ASSERT_EQUALS(last_cell.value<std::string>(), "C2");

        for (const auto row : rows)
        {
            for (const auto cell : row)
            {
                TS_ASSERT_EQUALS(cell.value<std::string>(), cell.reference().to_string());
            }
        }
    }
    
    void test_const_reverse_iterators()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();

        ws.append({"A1", "B1", "C1"});
        ws.append({"A2", "B2", "C2"});

        const xlnt::worksheet ws_const = ws;
        const auto rows = ws_const.rows();

        const auto first_row = *rows.rbegin();
        const auto first_cell = *first_row.rbegin();
        TS_ASSERT_EQUALS(first_cell.value<std::string>(), "C2");

        auto row_iter = rows.rend();
        row_iter--;
        const auto last_row = *row_iter;
        auto cell_iter = last_row.rend();
        cell_iter--;
        const auto last_cell = *cell_iter;
        TS_ASSERT_EQUALS(last_cell.value<std::string>(), "A1");

        for (auto ws_iter = rows.rbegin(); ws_iter != rows.rend(); ws_iter++)
        {
            const auto row = *ws_iter;

            for (auto row_iter = row.rbegin(); row_iter != row.rend(); row_iter++)
            {
                const auto cell = *row_iter;
                TS_ASSERT_EQUALS(cell.value<std::string>(), cell.reference().to_string());
            }
        }
    }
    
    void test_column_major_iterators()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();

        ws.append({"A1", "B1", "C1"});
        ws.append({"A2", "B2", "C2"});

        auto columns = ws.columns();

        auto first_column = *columns.begin();
        auto first_column_iter = first_column.begin();
        auto first_cell = *first_column_iter;
        TS_ASSERT_EQUALS(first_cell.value<std::string>(), "A1");
        first_column_iter++;
        auto second_cell = *first_column_iter;
        TS_ASSERT_EQUALS(second_cell.value<std::string>(), "A2");

        TS_ASSERT_EQUALS(first_cell, first_column.front());
        TS_ASSERT_EQUALS(second_cell, first_column.back());

        auto last_column = *(--columns.end());
        auto last_column_iter = last_column.end();
        last_column_iter--;
        auto last_cell = *last_column_iter;
        TS_ASSERT_EQUALS(last_cell.value<std::string>(), "C2");
        last_column_iter--;
        auto penultimate_cell = *last_column_iter;
        TS_ASSERT_EQUALS(penultimate_cell.value<std::string>(), "C1");

        for (auto column : columns)
        {
            for (auto cell : column)
            {
                TS_ASSERT_EQUALS(cell.value<std::string>(), cell.reference().to_string());
            }
        }
    }

    void test_reverse_column_major_iterators()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();

        ws.append({"A1", "B1", "C1"});
        ws.append({"A2", "B2", "C2"});

        auto columns = ws.columns();

        auto column_iter = columns.rbegin();
        auto first_column = *column_iter;

        auto first_column_iter = first_column.rbegin();
        auto first_cell = *first_column_iter;
        TS_ASSERT_EQUALS(first_cell.value<std::string>(), "C2");
        first_column_iter++;
        auto second_cell = *first_column_iter;
        TS_ASSERT_EQUALS(second_cell.value<std::string>(), "C1");

        auto last_column = *(--columns.rend());
        auto last_column_iter = last_column.rend();
        last_column_iter--;
        auto last_cell = *last_column_iter;
        TS_ASSERT_EQUALS(last_cell.value<std::string>(), "A1");
        last_column_iter--;
        auto penultimate_cell = *last_column_iter;
        TS_ASSERT_EQUALS(penultimate_cell.value<std::string>(), "A2");

        for (auto column_iter = columns.rbegin(); column_iter != columns.rend(); ++column_iter)
        {
            auto column = *column_iter;
            
            for (auto cell_iter = column.rbegin(); cell_iter != column.rend(); ++cell_iter)
            {
                auto cell = *cell_iter;

                TS_ASSERT_EQUALS(cell.value<std::string>(), cell.reference().to_string());
            }
        }
    }

    void test_const_column_major_iterators()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();

        ws.append({"A1", "B1", "C1"});
        ws.append({"A2", "B2", "C2"});

        const xlnt::worksheet ws_const = ws;
        const auto columns = ws_const.columns();

        const auto first_column = *columns.begin();
        auto first_column_iter = first_column.begin();
        const auto first_cell = *first_column_iter;
        TS_ASSERT_EQUALS(first_cell.value<std::string>(), "A1");
        first_column_iter++;
        const auto second_cell = *first_column_iter;
        TS_ASSERT_EQUALS(second_cell.value<std::string>(), "A2");

        const auto last_column = *(--columns.end());
        auto last_column_iter = last_column.end();
        last_column_iter--;
        const auto last_cell = *last_column_iter;
        TS_ASSERT_EQUALS(last_cell.value<std::string>(), "C2");
        last_column_iter--;
        const auto penultimate_cell = *last_column_iter;
        TS_ASSERT_EQUALS(penultimate_cell.value<std::string>(), "C1");

        for (const auto column : columns)
        {
            for (const auto cell : column)
            {
                TS_ASSERT_EQUALS(cell.value<std::string>(), cell.reference().to_string());
            }
        }
    }
    
    void test_const_reverse_column_major_iterators()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();

        ws.append({"A1", "B1", "C1"});
        ws.append({"A2", "B2", "C2"});

        const xlnt::worksheet ws_const = ws;
        const auto columns = ws_const.columns();

        const auto first_column = *columns.crbegin();
        auto first_column_iter = first_column.crbegin();
        const auto first_cell = *first_column_iter;
        TS_ASSERT_EQUALS(first_cell.value<std::string>(), "C2");
        first_column_iter++;
        const auto second_cell = *first_column_iter;
        TS_ASSERT_EQUALS(second_cell.value<std::string>(), "C1");

        const auto last_column = *(--columns.crend());
        auto last_column_iter = last_column.crend();
        last_column_iter--;
        const auto last_cell = *last_column_iter;
        TS_ASSERT_EQUALS(last_cell.value<std::string>(), "A1");
        last_column_iter--;
        const auto penultimate_cell = *last_column_iter;
        TS_ASSERT_EQUALS(penultimate_cell.value<std::string>(), "A2");

        for (auto column_iter = columns.crbegin(); column_iter != columns.crend(); ++column_iter)
        {
            const auto column = *column_iter;
            
            for (auto cell_iter = column.crbegin(); cell_iter != column.crend(); ++cell_iter)
            {
                const auto cell = *cell_iter;

                TS_ASSERT_EQUALS(cell.value<std::string>(), cell.reference().to_string());
            }
        }
    }
    
    void test_header()
    {
        xlnt::header_footer hf;
        using hf_loc = xlnt::header_footer::location;

        for (auto location : { hf_loc::left, hf_loc::center, hf_loc::right })
        {
            TS_ASSERT(!hf.has_header(location));
            TS_ASSERT(!hf.has_odd_even_header(location));
            TS_ASSERT(!hf.has_first_page_header(location));
            
            hf.header(location, "abc");

            TS_ASSERT(hf.has_header(location));
            TS_ASSERT(!hf.has_odd_even_header(location));
            TS_ASSERT(!hf.has_first_page_header(location));

            TS_ASSERT_EQUALS(hf.header(location), "abc");

            hf.clear_header(location);

            TS_ASSERT(!hf.has_header(location));
        }
    }
    
    void test_footer()
    {
        xlnt::header_footer hf;
        using hf_loc = xlnt::header_footer::location;

        for (auto location : { hf_loc::left, hf_loc::center, hf_loc::right })
        {
            TS_ASSERT(!hf.has_footer(location));
            TS_ASSERT(!hf.has_odd_even_footer(location));
            TS_ASSERT(!hf.has_first_page_footer(location));
            
            hf.footer(location, "abc");

            TS_ASSERT(hf.has_footer(location));
            TS_ASSERT(!hf.has_odd_even_footer(location));
            TS_ASSERT(!hf.has_first_page_footer(location));

            TS_ASSERT_EQUALS(hf.footer(location), "abc");

            hf.clear_footer(location);

            TS_ASSERT(!hf.has_footer(location));
        }
    }
    
    void test_page_setup()
    {
		xlnt::page_setup setup;
		setup.page_break(xlnt::page_break::column);
        TS_ASSERT_EQUALS(setup.page_break(), xlnt::page_break::column);
		setup.scale(1.23);
        TS_ASSERT_EQUALS(setup.scale(), 1.23);
    }

    void test_unique_sheet_name()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();

        auto active_name = ws.title();
        auto next_name = ws.unique_sheet_name(active_name);

        TS_ASSERT_DIFFERS(active_name, next_name);
    }

    void test_page_margins()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        auto margins = ws.page_margins();

        margins.top(0);
        margins.bottom(1);
        margins.header(2);
        margins.footer(3);
        margins.left(4);
        margins.right(5);
    }

    void test_garbage_collect()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();

        auto dimensions = ws.calculate_dimension();
        TS_ASSERT_EQUALS(dimensions, xlnt::range_reference("A1", "A1"));

        ws.cell("B2").value("text");
        ws.garbage_collect();

        dimensions = ws.calculate_dimension();
        TS_ASSERT_EQUALS(dimensions, xlnt::range_reference("B2", "B2"));
    }

    void test_has_cell()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();

        ws.cell("A3").value("test");

        TS_ASSERT(!ws.has_cell("A2"));
        TS_ASSERT(ws.has_cell("A3"));
    }

    void test_get_range_by_string()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();

        ws.cell("A2").value(3.14);
        ws.cell("A3").value(true);
        ws.cell("B2").value("text");
        ws.cell("B3").value(false);

        auto range = ws.range("A2:B3");
        auto range_iter = range.begin();
        auto row = *range_iter;
        auto row_iter = row.begin();
        TS_ASSERT_EQUALS((*row_iter).value<double>(), 3.14);
        TS_ASSERT_EQUALS((*row_iter).reference(), "A2");
        TS_ASSERT_EQUALS((*row_iter), row.front());

        row_iter++;
        TS_ASSERT_EQUALS((*row_iter).value<std::string>(), "text");
        TS_ASSERT_EQUALS((*row_iter).reference(), "B2");
        TS_ASSERT_EQUALS((*row_iter), row.back());

        range_iter++;
        row = *range_iter;
        row_iter = row.begin();
        TS_ASSERT_EQUALS((*row_iter).value<bool>(), true);
        TS_ASSERT_EQUALS((*row_iter).reference(), "A3");

        range_iter = range.end();
        range_iter--;
        row = *range_iter;
        row_iter = row.end();
        row_iter--;
        TS_ASSERT_EQUALS((*row_iter).value<bool>(), false);
    }

    void test_operators()
    {
        xlnt::workbook wb;

        wb.create_sheet();
        wb.create_sheet();

        auto ws1 = wb[1];
        auto ws2 = wb[2];

        TS_ASSERT_DIFFERS(ws1, ws2);

        ws1[xlnt::cell_reference("A2")].value(true);

        const auto ws1_const = ws1;
        TS_ASSERT_EQUALS(ws1[xlnt::cell_reference("A2")].value<bool>(), true);
        TS_ASSERT_EQUALS((*(*ws1[xlnt::range_reference("A2:A2")].begin()).begin()).value<bool>(), true);

        ws1.create_named_range("rangey", "A2:A2");
        TS_ASSERT_EQUALS(ws1[std::string("rangey")], ws1.range("A2:A2"));
        TS_ASSERT_EQUALS(ws1[std::string("A2:A2")], ws1.range("A2:A2"));
        TS_ASSERT(ws1[std::string("rangey")] != ws1.range("A2:A3"));

        TS_ASSERT_EQUALS(ws1[std::string("rangey")].cell("A1"), ws1.cell("A2"));
    }

    void test_reserve()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();

        ws.reserve(1000);
        //TODO: actual tests go here
    }

    void test_iterate()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();

        ws.cell("B3").value("B3");
        ws.cell("C7").value("C7");

        for (auto row : ws)
        {
            for (auto cell : row)
            {
                if (cell.has_value())
                {
                    TS_ASSERT_EQUALS(cell.reference().to_string(), cell.value<std::string>());
                }
            }
        }

        const auto ws_const = ws;

        for (auto row : ws_const)
        {
            for (auto cell : row)
            {
                if (cell.has_value())
                {
                    TS_ASSERT_EQUALS(cell.reference().to_string(), cell.value<std::string>());
                }
            }
        }

        const auto const_range = ws_const.range("B3:C7");
        auto const_range_iter = const_range.cbegin();
        const_range_iter++;
        const_range_iter--;
        TS_ASSERT_EQUALS(const_range_iter, const_range.begin());
    }

    void test_range_reference()
    {
        xlnt::range_reference ref1("A1:A1");
        TS_ASSERT(ref1.is_single_cell());
        xlnt::range_reference ref2("A1:B2");

        TS_ASSERT(!ref2.is_single_cell());
        TS_ASSERT(ref1 == xlnt::range_reference("A1:A1"));
        TS_ASSERT(ref1 != ref2);
        TS_ASSERT(ref1 == "A1:A1");
        TS_ASSERT(ref1 == std::string("A1:A1"));
        TS_ASSERT(std::string("A1:A1") == ref1);
        TS_ASSERT("A1:A1" == ref1);
        TS_ASSERT(ref1 != "A1:B2");
        TS_ASSERT(ref1 != std::string("A1:B2"));
        TS_ASSERT(std::string("A1:B2") != ref1);
        TS_ASSERT("A1:B2" != ref1);
    }

    void test_get_point_pos()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();

        TS_ASSERT_EQUALS(ws.point_pos({0, 0}), "A1");
    }

	void test_named_range_named_cell_reference()
	{
		xlnt::workbook wb;
		auto ws = wb.active_sheet();
		TS_ASSERT_THROWS(ws.create_named_range("A1", "A2"), xlnt::invalid_parameter);
		TS_ASSERT_THROWS(ws.create_named_range("XFD1048576", "A2"), xlnt::invalid_parameter);
		TS_ASSERT_THROWS_NOTHING(ws.create_named_range("XFE1048576", "A2"));
		TS_ASSERT_THROWS_NOTHING(ws.create_named_range("XFD1048577", "A2"));
	}

    void test_iteration_skip_empty()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        ws.cell("A1").value("A1");
        ws.cell("F6").value("F6");

        {
            std::vector<xlnt::cell> cells;

            for (auto row : ws)
            {
                for (auto cell : row)
                {
                    cells.push_back(cell);
                }
            }

            TS_ASSERT_EQUALS(cells.size(), 2);
            TS_ASSERT_EQUALS(cells[0].value<std::string>(), "A1");
            TS_ASSERT_EQUALS(cells[1].value<std::string>(), "F6");
        }

        const auto ws_const = ws;

        {
            std::vector<xlnt::cell> cells;

            for (auto row : ws_const)
            {
                for (auto cell : row)
                {
                    cells.push_back(cell);
                }
            }

            TS_ASSERT_EQUALS(cells.size(), 2);
            TS_ASSERT_EQUALS(cells[0].value<std::string>(), "A1");
            TS_ASSERT_EQUALS(cells[1].value<std::string>(), "F6");
        }
    }
};
