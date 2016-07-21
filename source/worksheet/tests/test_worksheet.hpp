#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/worksheet/footer.hpp>
#include <xlnt/worksheet/header.hpp>
#include <xlnt/worksheet/header_footer.hpp>
#include <xlnt/worksheet/worksheet.hpp>

class test_worksheet : public CxxTest::TestSuite
{
public:
    void test_new_worksheet()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        TS_ASSERT(ws.get_workbook() == wb);
    }
    
    void test_get_cell()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        auto cell = ws.get_cell(xlnt::cell_reference(1, 1));
        TS_ASSERT_EQUALS(cell.get_reference(), "A1");
    }
    
    void test_invalid_cell()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        TS_ASSERT_THROWS(xlnt::cell_reference invalid(xlnt::column_t((xlnt::column_t::index_t)0), 0), xlnt::value_error);
    }
    
    void test_worksheet_dimension()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        
        TS_ASSERT_EQUALS("A1:A1", ws.calculate_dimension());
        ws.get_cell("B12").set_value("AAA");
        TS_ASSERT_EQUALS("B12:B12", ws.calculate_dimension());
    }
    
    void test_squared_range()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        
        const std::vector<std::vector<std::string>> expected =
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
    
    void test_fill_rows()
    {
        std::size_t row = 0;
        std::size_t column = 0;
        std::string coordinate = "A1";
        
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        
        ws.get_cell("A1").set_value("first");
        ws.get_cell("C9").set_value("last");
        
        TS_ASSERT_EQUALS(ws.calculate_dimension(), "A1:C9");
        TS_ASSERT_EQUALS(ws.rows()[row][column].get_reference(), coordinate);
        
        row = 8;
        column = 2;
        coordinate = "C9";
        
        TS_ASSERT_EQUALS(ws.rows()[row][column].get_reference(), coordinate);
    }
    
    void test_iter_rows()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        
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
                TS_ASSERT_EQUALS(cell.get_reference(), *expected_cell_iter);
                expected_cell_iter++;
            }
            
            expected_row_iter++;
        }
    }
    
    void test_iter_rows_offset()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
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
                TS_ASSERT_EQUALS(cell.get_reference(), *expected_cell_iter);
                expected_cell_iter++;
            }
            
            expected_row_iter++;
        }
    }

    void test_iter_rows_offset_int_int()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
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
                TS_ASSERT_EQUALS(cell.get_reference(), *expected_cell_iter);
                expected_cell_iter++;
            }

            expected_row_iter++;
        }
    }
    
    void test_get_named_range()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        wb.create_named_range("test_range", ws, "C5");
        auto xlrange = ws.get_named_range("test_range");
        TS_ASSERT_EQUALS(1, xlrange.length());
        TS_ASSERT_EQUALS(1, xlrange[0].num_cells());
        TS_ASSERT_EQUALS(5, xlrange[0][0].get_row());

        ws.create_named_range("test_range2", "C6");
        auto xlrange2 = ws.get_named_range("test_range2");
        TS_ASSERT_EQUALS(1, xlrange2.length());
        TS_ASSERT_EQUALS(1, xlrange2[0].num_cells());
        TS_ASSERT_EQUALS(6, xlrange2[0][0].get_row());
    }
    
    void test_get_bad_named_range()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        TS_ASSERT_THROWS(ws.get_named_range("bad_range"), xlnt::key_error);
    }
    
    void test_get_named_range_wrong_sheet()
    {
        xlnt::workbook wb;
        wb.create_sheet();
        wb.create_sheet();
        auto ws1 = wb[0];
        auto ws2 = wb[1];
        wb.create_named_range("wrong_sheet_range", ws1, "C5");
        TS_ASSERT_THROWS(ws2.get_named_range("wrong_sheet_range"), xlnt::named_range_exception);
    }
    
    void test_remove_named_range_bad()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        TS_ASSERT_THROWS(ws.remove_named_range("bad_range"), std::runtime_error);
    }

    void test_cell_alternate_coordinates()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        auto cell = ws.get_cell(xlnt::cell_reference(4, 8));
        TS_ASSERT_EQUALS(cell.get_reference(), "D8");
    }
    
    // void test_cell_insufficient_coordinates() {}
    
    void test_cell_range_name()
    {
        xlnt::workbook wb;
        auto ws = wb.get_active_sheet();
        wb.create_named_range("test_range_single", ws, "B12");
        auto c_range_name = ws.get_named_range("test_range_single");
        auto c_range_coord = ws.get_range("B12");
        auto c_cell = ws.get_cell("B12");
        TS_ASSERT_EQUALS(c_range_coord, c_range_name);
        TS_ASSERT(c_range_coord[0][0] == c_cell);
    }
    
    void test_hyperlink_value()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        ws.get_cell("A1").set_hyperlink("http://test.com");
        TS_ASSERT_EQUALS("http://test.com", ws.get_cell("A1").get_value<std::string>());
        ws.get_cell("A1").set_value("test");
        TS_ASSERT_EQUALS("test", ws.get_cell("A1").get_value<std::string>());
        TS_ASSERT_EQUALS(ws.get_cell("A1").get_hyperlink().get_target_uri(), "http://test.com");
    }
    
    void test_append()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        ws.append(std::vector<std::string> {"value"});
        TS_ASSERT_EQUALS("value", ws.get_cell("A1").get_value<std::string>());
    }
    
    void test_append_list()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        
        ws.append(std::vector<std::string> {"This is A1", "This is B1"});
        
        TS_ASSERT_EQUALS("This is A1", ws.get_cell("A1").get_value<std::string>());
        TS_ASSERT_EQUALS("This is B1", ws.get_cell("B1").get_value<std::string>());
    }
    
    void test_append_dict_letter()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        
        const std::unordered_map<std::string, std::string> dict_letter =
        {
            { "A", "This is A1" },
            { "C", "This is C1" }
        };
        
        ws.append(dict_letter);
        
        TS_ASSERT_EQUALS("This is A1", ws.get_cell("A1").get_value<std::string>());
        TS_ASSERT_EQUALS("This is C1", ws.get_cell("C1").get_value<std::string>());
    }
    
    void test_append_dict_index()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        
        const std::unordered_map<int, std::string> dict_index =
        {
            { 1, "This is A1" },
            { 3, "This is C1" }
        };
        
        ws.append(dict_index);
        
        TS_ASSERT_EQUALS("This is A1", ws.get_cell("A1").get_value<std::string>());
        TS_ASSERT_EQUALS("This is C1", ws.get_cell("C1").get_value<std::string>());
    }
    
    // void test_bad_append() {}
    
    // void test_append_range() {}
    
    void test_append_iterator()
    {
        std::vector<int> range;
        
        for(int i = 0; i < 30; i++)
        {
            range.push_back(i);
        }
        
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        ws.append(range.begin(), range.end());
        
        TS_ASSERT_EQUALS(ws[xlnt::cell_reference("AD1")].get_value<int>(), 29);
    }
    
    void test_append_2d_list()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        
        ws.append(std::vector<std::string> {"This is A1", "This is B1"});
        ws.append(std::vector<std::string> {"This is A2", "This is B2"});
        
        auto vals = ws.get_range("A1:B2");
        
        TS_ASSERT_EQUALS(vals[0][0].get_value<std::string>(), "This is A1");
        TS_ASSERT_EQUALS(vals[0][1].get_value<std::string>(), "This is B1");
        TS_ASSERT_EQUALS(vals[1][0].get_value<std::string>(), "This is A2");
        TS_ASSERT_EQUALS(vals[1][1].get_value<std::string>(), "This is B2");
    }
    
    // void test_append_cell()
    // {
        // Right now, a cell cannot be created without a parent worksheet.
        // This should be possible by instantiating a cell_impl, but the
        // memory management is complicated.
        // xlnt::workbook wb;
        // xlnt::worksheet ws(wb);
        // xlnt::cell cell;
        // cell.set_value(25);
        
        // ws.append({ cell });
        
        // TS_ASSERT_EQUALS(cell.get_value<int>(), 25);
    // }
    
    void test_rows()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        
        ws.get_cell("A1").set_value("first");
        ws.get_cell("C9").set_value("last");
        
        auto rows = ws.rows();
        
        TS_ASSERT_EQUALS(rows.length(), 9);
        
        auto first_row = rows[0];
        auto last_row = rows[8];
        
        TS_ASSERT_EQUALS(first_row[0].get_value<std::string>(), "first");
        TS_ASSERT_EQUALS(first_row[0].get_reference(), "A1");
        TS_ASSERT_EQUALS(last_row[2].get_value<std::string>(), "last");
    }
    
    void test_no_rows()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        
        TS_ASSERT_EQUALS(ws.rows().length(), 1);
        TS_ASSERT_EQUALS(ws.rows()[0].length(), 1);
    }
    
    void test_no_cols()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        
        TS_ASSERT_EQUALS(ws.columns().length(), 1);
        TS_ASSERT_EQUALS(ws.columns()[0].length(), 1);
    }
    
    void test_one_cell()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        
        auto cell = ws.get_cell("A1");
        
        TS_ASSERT_EQUALS(ws.columns().length(), 1);
        TS_ASSERT_EQUALS(ws.columns()[0].length(), 1);
        TS_ASSERT_EQUALS(ws.columns()[0][0], cell);
    }
    
    // void test_by_col() {}
    
    void test_cols()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        
        ws.get_cell("A1").set_value("first");
        ws.get_cell("C9").set_value("last");
        
        auto cols = ws.columns();
        
        TS_ASSERT_EQUALS(cols.length(), 3);
        
        TS_ASSERT_EQUALS(cols[0][0].get_value<std::string>(), "first");
        TS_ASSERT_EQUALS(cols[2][8].get_value<std::string>(), "last");
    }
    
    void test_auto_filter()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        
        ws.auto_filter(ws.get_range("a1:f1"));
        TS_ASSERT_EQUALS(ws.get_auto_filter(), "A1:F1");
        
        ws.unset_auto_filter();
        TS_ASSERT_EQUALS(ws.has_auto_filter(), false);
        
        ws.auto_filter("c1:g9");
        TS_ASSERT_EQUALS(ws.get_auto_filter(), "C1:G9");
    }
    
    void test_getitem()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        xlnt::cell cell = ws[xlnt::cell_reference("A1")];
        TS_ASSERT_EQUALS(cell.get_reference().to_string(), "A1");
        TS_ASSERT_EQUALS(cell.get_data_type(), xlnt::cell::type::null);
    }
    
    // void test_getitem_invalid() {}
    
    void test_setitem()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        ws[xlnt::cell_reference("A12")].set_value(5);
        TS_ASSERT(ws[xlnt::cell_reference("A12")].get_value<int>() == 5);
    }
    
    void test_getslice()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        auto cell_range = ws("A1", "B2");
        TS_ASSERT_EQUALS(cell_range[0][0], ws.get_cell("A1"));
        TS_ASSERT_EQUALS(cell_range[1][0], ws.get_cell("A2"));
        TS_ASSERT_EQUALS(cell_range[0][1], ws.get_cell("B1"));
        TS_ASSERT_EQUALS(cell_range[1][1], ws.get_cell("B2"));
    }
    
    // void test_get_column() {}
    // void test_get_row() {}
    
    void test_freeze()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        
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
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        ws.get_cell("A2").set_value("test");
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
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        TS_ASSERT_EQUALS(ws.get_merged_ranges().size(), 0);
    }
    
    void test_merge_range_string()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        ws.get_cell("A1").set_value(1);
        ws.get_cell("D4").set_value(16);
        ws.merge_cells("A1:D4");
        std::vector<xlnt::range_reference> expected = { xlnt::range_reference("A1:D4") };
        TS_ASSERT_EQUALS(ws.get_merged_ranges(), expected);
        TS_ASSERT(!ws.get_cell("D4").has_value());
    }
    
    void test_merge_coordinate()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        ws.merge_cells(1, 1, 4, 4);
        std::vector<xlnt::range_reference> expected = { xlnt::range_reference("A1:D4") };
        TS_ASSERT_EQUALS(ws.get_merged_ranges(), expected);
    }
    
    void test_unmerge_bad()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);

        TS_ASSERT_THROWS(ws.unmerge_cells("A1:D3"), std::runtime_error);
    }

    void test_unmerge_range_string()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        ws.merge_cells("A1:D4");
        TS_ASSERT_EQUALS(ws.get_merged_ranges().size(), 1);
        ws.unmerge_cells("A1:D4");
        TS_ASSERT_EQUALS(ws.get_merged_ranges().size(), 0);
    }

    void test_unmerge_coordinate()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        ws.merge_cells("A1:D4");
        TS_ASSERT_EQUALS(ws.get_merged_ranges().size(), 1);
        ws.unmerge_cells(1, 1, 4, 4);
        TS_ASSERT_EQUALS(ws.get_merged_ranges().size(), 0);
    }

    void test_print_titles_old()
    {
        xlnt::workbook wb;
        
        {
            wb.clear();
            auto ws = wb.create_sheet();
            ws.add_print_title(3);
            TS_ASSERT_EQUALS(ws.get_print_titles(), "Sheet!1:3");
        }
        
        {
            wb.clear();
            auto ws = wb.create_sheet();
            ws.add_print_title(4, "cols");
            TS_ASSERT_EQUALS(ws.get_print_titles(), "Sheet!A:D");
        }
    }
    
    void test_print_titles_new()
    {
        xlnt::workbook wb;
        
        {
            wb.clear();
            auto ws = wb.create_sheet();
            ws.set_print_title_rows("1:4");
            TS_ASSERT_EQUALS(ws.get_print_titles(), "Sheet!1:4");
        }
        
        {
            wb.clear();
            auto ws = wb.create_sheet();
            ws.set_print_title_cols("A:F");
            TS_ASSERT_EQUALS(ws.get_print_titles(), "Sheet!A:F");
        }

        {
            wb.clear();
            auto ws = wb.create_sheet();
            ws.set_print_title_rows("1:2");
            ws.set_print_title_cols("C:D");
            TS_ASSERT_EQUALS(ws.get_print_titles(), "Sheet!1:2,Sheet!C:D");
        }
    }
    
    void test_print_area()
    {
        xlnt::workbook wb;
        auto ws = wb.get_active_sheet();
        ws.set_print_area("A1:F5");
        TS_ASSERT_EQUALS(ws.get_print_area(), "$A$1:$F$5");
    }
    
    void test_freeze_panes_horiz()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        ws.freeze_panes("A4");
        
        auto view = ws.get_sheet_view();
        TS_ASSERT_EQUALS(view.get_selections().size(), 1);
        TS_ASSERT_EQUALS(view.get_selections()[0].get_active_cell(), "A1");
        TS_ASSERT_EQUALS(view.get_selections()[0].get_pane(), xlnt::pane_corner::bottom_left);
        TS_ASSERT_EQUALS(view.get_selections()[0].get_sqref(), "A1");
        TS_ASSERT_EQUALS(view.get_pane().active_pane, xlnt::pane_corner::bottom_left);
        TS_ASSERT_EQUALS(view.get_pane().state, xlnt::pane_state::frozen);
        TS_ASSERT_EQUALS(view.get_pane().top_left_cell, "A4");
        TS_ASSERT_EQUALS(view.get_pane().y_split, 3);
    }
    
    void test_freeze_panes_vert()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        ws.freeze_panes("D1");
        
        auto view = ws.get_sheet_view();
        TS_ASSERT_EQUALS(view.get_selections().size(), 1);
        TS_ASSERT_EQUALS(view.get_selections()[0].get_active_cell(), "A1");
        TS_ASSERT_EQUALS(view.get_selections()[0].get_pane(), xlnt::pane_corner::top_right);
        TS_ASSERT_EQUALS(view.get_selections()[0].get_sqref(), "A1");
        TS_ASSERT_EQUALS(view.get_pane().active_pane, xlnt::pane_corner::top_right);
        TS_ASSERT_EQUALS(view.get_pane().state, xlnt::pane_state::frozen);
        TS_ASSERT_EQUALS(view.get_pane().top_left_cell, "D1");
        TS_ASSERT_EQUALS(view.get_pane().x_split, 3);
    }
    
    void test_freeze_panes_both()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        ws.freeze_panes("D4");
        
        auto view = ws.get_sheet_view();
        TS_ASSERT_EQUALS(view.get_selections().size(), 3);
        TS_ASSERT_EQUALS(view.get_selections()[0].get_pane(), xlnt::pane_corner::top_right);
        TS_ASSERT_EQUALS(view.get_selections()[1].get_pane(), xlnt::pane_corner::bottom_left);
        TS_ASSERT_EQUALS(view.get_selections()[2].get_active_cell(), "A1");
        TS_ASSERT_EQUALS(view.get_selections()[2].get_pane(), xlnt::pane_corner::bottom_right);
        TS_ASSERT_EQUALS(view.get_selections()[2].get_sqref(), "A1");
        TS_ASSERT_EQUALS(view.get_pane().active_pane, xlnt::pane_corner::bottom_right);
        TS_ASSERT_EQUALS(view.get_pane().state, xlnt::pane_state::frozen);
        TS_ASSERT_EQUALS(view.get_pane().top_left_cell, "D4");
        TS_ASSERT_EQUALS(view.get_pane().x_split, 3);
        TS_ASSERT_EQUALS(view.get_pane().y_split, 3);
    }
    
    void test_min_column()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        TS_ASSERT_EQUALS(ws.get_lowest_column(), 1);
    }
    
    void test_max_column()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        ws[xlnt::cell_reference("F1")].set_value(10);
        ws[xlnt::cell_reference("F2")].set_value(32);
        ws[xlnt::cell_reference("F3")].set_formula("=F1+F2");
        ws[xlnt::cell_reference("A4")].set_formula("=A1+A2+A3");
        TS_ASSERT_EQUALS(ws.get_highest_column(), 6);
    }

    void test_min_row()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        TS_ASSERT_EQUALS(ws.get_lowest_row(), 1);
    }
    
    void test_max_row()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);
        ws.append();
        ws.append(std::vector<int> { 5 });
        ws.append();
        ws.append(std::vector<int> { 4 });
        TS_ASSERT_EQUALS(ws.get_highest_row(), 4);
    }
    
    void test_const_iterators()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);

        ws.append({"A1", "B1", "C1"});
        ws.append({"A2", "B2", "C2"});

        const xlnt::worksheet ws_const = ws;
        const auto rows = ws_const.rows();

        const auto first_row = *rows.begin();
        const auto first_cell = *first_row.begin();
        TS_ASSERT_EQUALS(first_cell.get_value<std::string>(), "A1");

        const auto last_row = *(--rows.end());
        const auto last_cell = *(--last_row.end());
        TS_ASSERT_EQUALS(last_cell.get_value<std::string>(), "C2");

        for (const auto row : rows)
        {
            for (const auto cell : row)
            {
                TS_ASSERT_EQUALS(cell.get_value<std::string>(), cell.get_reference().to_string());
            }
        }
    }
    
    void test_const_reverse_iterators()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);

        ws.append({"A1", "B1", "C1"});
        ws.append({"A2", "B2", "C2"});

        const xlnt::worksheet ws_const = ws;
        const auto rows = ws_const.rows();

        const auto first_row = *rows.rbegin();
        const auto first_cell = *first_row.rbegin();
        TS_ASSERT_EQUALS(first_cell.get_value<std::string>(), "C2");

        auto row_iter = rows.rend();
        row_iter--;
        const auto last_row = *row_iter;
        auto cell_iter = last_row.rend();
        cell_iter--;
        const auto last_cell = *cell_iter;
        TS_ASSERT_EQUALS(last_cell.get_value<std::string>(), "A1");

        for (auto ws_iter = rows.rbegin(); ws_iter != rows.rend(); ws_iter++)
        {
            const auto row = *ws_iter;

            for (auto row_iter = row.rbegin(); row_iter != row.rend(); row_iter++)
            {
                const auto cell = *row_iter;
                TS_ASSERT_EQUALS(cell.get_value<std::string>(), cell.get_reference().to_string());
            }
        }
    }
    
    void test_column_major_iterators()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);

        ws.append({"A1", "B1", "C1"});
        ws.append({"A2", "B2", "C2"});

        auto columns = ws.columns();

        auto first_column = *columns.begin();
        auto first_column_iter = first_column.begin();
        auto first_cell = *first_column_iter;
        TS_ASSERT_EQUALS(first_cell.get_value<std::string>(), "A1");
        first_column_iter++;
        auto second_cell = *first_column_iter;
        TS_ASSERT_EQUALS(second_cell.get_value<std::string>(), "A2");

        auto last_column = *(--columns.end());
        auto last_column_iter = last_column.end();
        last_column_iter--;
        auto last_cell = *last_column_iter;
        TS_ASSERT_EQUALS(last_cell.get_value<std::string>(), "C2");
        last_column_iter--;
        auto penultimate_cell = *last_column_iter;
        TS_ASSERT_EQUALS(penultimate_cell.get_value<std::string>(), "C1");

        for (auto column : columns)
        {
            for (auto cell : column)
            {
                TS_ASSERT_EQUALS(cell.get_value<std::string>(), cell.get_reference().to_string());
            }
        }
    }

    void test_reverse_column_major_iterators()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);

        ws.append({"A1", "B1", "C1"});
        ws.append({"A2", "B2", "C2"});

        auto columns = ws.columns();

        auto column_iter = columns.rbegin();
        auto first_column = *column_iter;

        auto first_column_iter = first_column.rbegin();
        auto first_cell = *first_column_iter;
        TS_ASSERT_EQUALS(first_cell.get_value<std::string>(), "C2");
        first_column_iter++;
        auto second_cell = *first_column_iter;
        TS_ASSERT_EQUALS(second_cell.get_value<std::string>(), "C1");

        auto last_column = *(--columns.rend());
        auto last_column_iter = last_column.rend();
        last_column_iter--;
        auto last_cell = *last_column_iter;
        TS_ASSERT_EQUALS(last_cell.get_value<std::string>(), "A1");
        last_column_iter--;
        auto penultimate_cell = *last_column_iter;
        TS_ASSERT_EQUALS(penultimate_cell.get_value<std::string>(), "A2");

        for (auto column_iter = columns.rbegin(); column_iter != columns.rend(); ++column_iter)
        {
            auto column = *column_iter;
            
            for (auto cell_iter = column.rbegin(); cell_iter != column.rend(); ++cell_iter)
            {
                auto cell = *cell_iter;

                TS_ASSERT_EQUALS(cell.get_value<std::string>(), cell.get_reference().to_string());
            }
        }
    }

    void test_const_column_major_iterators()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);

        ws.append({"A1", "B1", "C1"});
        ws.append({"A2", "B2", "C2"});

        const xlnt::worksheet ws_const = ws;
        const auto columns = ws_const.columns();

        const auto first_column = *columns.begin();
        auto first_column_iter = first_column.begin();
        const auto first_cell = *first_column_iter;
        TS_ASSERT_EQUALS(first_cell.get_value<std::string>(), "A1");
        first_column_iter++;
        const auto second_cell = *first_column_iter;
        TS_ASSERT_EQUALS(second_cell.get_value<std::string>(), "A2");

        const auto last_column = *(--columns.end());
        auto last_column_iter = last_column.end();
        last_column_iter--;
        const auto last_cell = *last_column_iter;
        TS_ASSERT_EQUALS(last_cell.get_value<std::string>(), "C2");
        last_column_iter--;
        const auto penultimate_cell = *last_column_iter;
        TS_ASSERT_EQUALS(penultimate_cell.get_value<std::string>(), "C1");

        for (const auto column : columns)
        {
            for (const auto cell : column)
            {
                TS_ASSERT_EQUALS(cell.get_value<std::string>(), cell.get_reference().to_string());
            }
        }
    }
    
    void test_const_reverse_column_major_iterators()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);

        ws.append({"A1", "B1", "C1"});
        ws.append({"A2", "B2", "C2"});

        const xlnt::worksheet ws_const = ws;
        const auto columns = ws_const.columns();

        const auto first_column = *columns.crbegin();
        auto first_column_iter = first_column.crbegin();
        const auto first_cell = *first_column_iter;
        TS_ASSERT_EQUALS(first_cell.get_value<std::string>(), "C2");
        first_column_iter++;
        const auto second_cell = *first_column_iter;
        TS_ASSERT_EQUALS(second_cell.get_value<std::string>(), "C1");

        const auto last_column = *(--columns.crend());
        auto last_column_iter = last_column.crend();
        last_column_iter--;
        const auto last_cell = *last_column_iter;
        TS_ASSERT_EQUALS(last_cell.get_value<std::string>(), "A1");
        last_column_iter--;
        const auto penultimate_cell = *last_column_iter;
        TS_ASSERT_EQUALS(penultimate_cell.get_value<std::string>(), "A2");

        for (auto column_iter = columns.crbegin(); column_iter != columns.crend(); ++column_iter)
        {
            const auto column = *column_iter;
            
            for (auto cell_iter = column.crbegin(); cell_iter != column.crend(); ++cell_iter)
            {
                const auto cell = *cell_iter;

                TS_ASSERT_EQUALS(cell.get_value<std::string>(), cell.get_reference().to_string());
            }
        }
    }
    
    void test_header()
    {
        xlnt::workbook wb;
        auto ws = wb.get_active_sheet();

        auto &header_footer = ws.get_header_footer();

        for (auto header_pointer : { &header_footer.get_left_header(),
            &header_footer.get_center_header(), &header_footer.get_right_header() })
        {
            auto &header = *header_pointer;

            TS_ASSERT(header.is_default());
            header.set_text("abc");
            header.set_font_name("def");
            header.set_font_size(121);
            header.set_font_color("ghi");
            TS_ASSERT(!header.is_default());
        }
    }
    
    void test_footer()
    {
        xlnt::workbook wb;
        auto ws = wb.get_active_sheet();

        auto &header_footer = ws.get_header_footer();

        for (auto header_pointer : { &header_footer.get_left_footer(),
            &header_footer.get_center_footer(), &header_footer.get_right_footer() })
        {
            auto &header = *header_pointer;

            TS_ASSERT(header.is_default());
            header.set_text("abc");
            header.set_font_name("def");
            header.set_font_size(121);
            header.set_font_color("ghi");
            TS_ASSERT(!header.is_default());
        }
    }
    
    void test_page_setup()
    {
        xlnt::workbook wb;
        auto ws = wb.get_active_sheet();
        TS_ASSERT(ws.get_page_setup().is_default());
        ws.get_page_setup().set_break(xlnt::page_break::column);
        TS_ASSERT_EQUALS(ws.get_page_setup().get_break(), xlnt::page_break::column);
        TS_ASSERT(!ws.get_page_setup().is_default());
        ws.get_page_setup().set_scale(1.23);
        TS_ASSERT_EQUALS(ws.get_page_setup().get_scale(), 1.23);
        TS_ASSERT(!ws.get_page_setup().is_default());
    }

    void test_unique_sheet_name()
    {
        xlnt::workbook wb;
        auto ws = wb.get_active_sheet();

        auto active_name = ws.get_title();
        auto next_name = ws.unique_sheet_name(active_name);

        TS_ASSERT_DIFFERS(active_name, next_name);
    }

    void test_page_margins()
    {
        xlnt::workbook wb;
        auto ws = wb.get_active_sheet();
        auto &margins = ws.get_page_margins();

        TS_ASSERT(margins.is_default());

        margins.set_top(0);
        margins.set_bottom(1);
        margins.set_header(2);
        margins.set_footer(3);
        margins.set_left(4);
        margins.set_right(5);
        TS_ASSERT(!margins.is_default());
    }

    void test_to_string()
    {
        xlnt::workbook wb;
        auto ws = wb.get_active_sheet();
        auto ws_string = ws.to_string();
        TS_ASSERT_EQUALS(ws_string, "<Worksheet \"Sheet\">");
    }

    void test_garbage_collect()
    {
        xlnt::workbook wb;
        auto ws = wb.get_active_sheet();

        auto dimensions = ws.calculate_dimension();
        TS_ASSERT_EQUALS(dimensions, xlnt::range_reference("A1", "A1"));

        ws.get_cell("B2").set_value("text");
        ws.garbage_collect();

        dimensions = ws.calculate_dimension();
        TS_ASSERT_EQUALS(dimensions, xlnt::range_reference("B2", "B2"));
    }

    void test_get_title_bad()
    {
        xlnt::worksheet ws;
        TS_ASSERT_THROWS(ws.get_title(), std::runtime_error);
    }

    void test_has_cell()
    {
        xlnt::workbook wb;
        auto ws = wb.get_active_sheet();

        ws.get_cell("A3").set_value("test");

        TS_ASSERT(!ws.has_cell("A2"));
        TS_ASSERT(ws.has_cell("A3"));
    }

    void test_get_range_by_string()
    {
        xlnt::workbook wb;
        auto ws = wb.get_active_sheet();

        ws.get_cell("A2").set_value(3.14);
        ws.get_cell("A3").set_value(true);
        ws.get_cell("B2").set_value("text");
        ws.get_cell("B3").set_value(false);

        auto range = ws.get_range("A2:B3");
        auto range_iter = range.begin();
        auto row = *range_iter;
        auto row_iter = row.begin();
        auto cell = *row_iter;
        TS_ASSERT_EQUALS(cell.get_value<double>(), 3.14);
        TS_ASSERT_EQUALS(row.front().get_reference(), "A2");

        row_iter++;
        cell = *row_iter;
        TS_ASSERT_EQUALS(cell.get_value<std::string>(), "text");
        TS_ASSERT_EQUALS(row.back().get_reference(), "B2");

        range_iter = range.end();
        range_iter--;
        row = *range_iter;
        row_iter = row.end();
        row_iter--;
        cell = *row_iter;
        TS_ASSERT_EQUALS(cell.get_value<bool>(), false);
    }

    void test_get_squared_range()
    {
        xlnt::workbook wb;
        auto ws = wb.get_active_sheet();

        ws.get_cell("A2").set_value(3.14);
        ws.get_cell("B3").set_value(false);

        auto range = ws.get_squared_range(1, 2, 2, 3);
        TS_ASSERT_EQUALS((*(*range.begin()).begin()).get_value<double>(), 3.14);
        TS_ASSERT_EQUALS((*(--(*(--range.end())).end())).get_value<bool>(), false);
    }

    void test_operators()
    {
        xlnt::workbook wb;

        wb.create_sheet();
        wb.create_sheet();

        auto ws1 = wb[1];
        auto ws2 = wb[2];

        TS_ASSERT_DIFFERS(ws1, ws2);

        ws1[xlnt::cell_reference("A2")].set_value(true);

        const auto ws1_const = ws1;
        TS_ASSERT_EQUALS(ws1[xlnt::cell_reference("A2")].get_value<bool>(), true);
        TS_ASSERT_EQUALS((*(*ws1[xlnt::range_reference("A2:A2")].begin()).begin()).get_value<bool>(), true);

        ws1.create_named_range("rangey", "A2:A2");
        TS_ASSERT_EQUALS(ws1[std::string("rangey")], ws1.get_range("A2:A2"));
        TS_ASSERT_EQUALS(ws1[std::string("A2:A2")], ws1.get_range("A2:A2"));
        TS_ASSERT(ws1[std::string("rangey")] != ws1.get_range("A2:A3"));

        TS_ASSERT_EQUALS(ws1[std::string("rangey")].get_cell("A1"), ws1.get_cell("A2"));
    }

    void test_reserve()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);

        ws.reserve(1000);
        //TODO: actual tests go here
    }

    void test_iterate()
    {
        xlnt::workbook wb;
        xlnt::worksheet ws(wb);

        ws.get_cell("B3").set_value("B3");
        ws.get_cell("C7").set_value("C7");

        for (auto row : ws)
        {
            for (auto cell : row)
            {
                if (cell.has_value())
                {
                    TS_ASSERT_EQUALS(cell.get_reference().to_string(), cell.get_value<std::string>());
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
                    TS_ASSERT_EQUALS(cell.get_reference().to_string(), cell.get_value<std::string>());
                }
            }
        }
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
};
