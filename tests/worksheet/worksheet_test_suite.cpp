// Copyright (c) 2014-2018 Thomas Fussell
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, WRISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE
//
// @license: http://www.opensource.org/licenses/mit-license.php
// @author: see AUTHORS file

#include <iostream>

#include <xlnt/cell/cell.hpp>
#include <xlnt/cell/hyperlink.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/worksheet/column_properties.hpp>
#include <xlnt/worksheet/row_properties.hpp>
#include <xlnt/worksheet/range.hpp>
#include <xlnt/worksheet/worksheet.hpp>
#include <xlnt/worksheet/header_footer.hpp>
#include <xlnt/worksheet/worksheet.hpp>
#include <helpers/test_suite.hpp>

class worksheet_test_suite : public test_suite
{
public:
    worksheet_test_suite()
    {
        register_test(test_new_worksheet);
        register_test(test_cell);
        register_test(test_invalid_cell);
        register_test(test_worksheet_dimension);
        register_test(test_fill_rows);
        register_test(test_get_named_range);
        register_test(test_get_bad_named_range);
        register_test(test_get_named_range_wrong_sheet);
        register_test(test_remove_named_range_bad);
        register_test(test_cell_alternate_coordinates);
        register_test(test_cell_range_name);
        register_test(test_rows);
        register_test(test_no_rows);
        register_test(test_no_cols);
        register_test(test_one_cell);
        register_test(test_cols);
        register_test(test_auto_filter);
        register_test(test_getitem);
        register_test(test_setitem);
        register_test(test_getslice);
        register_test(test_freeze);
        register_test(test_merged_cells_lookup);
        register_test(test_merged_cell_ranges);
        register_test(test_merge_range_string);
        register_test(test_unmerge_bad);
        register_test(test_unmerge_range_string);
        register_test(test_print_titles_old);
        register_test(test_print_titles_new);
        register_test(test_print_area);
        register_test(test_freeze_panes_horiz);
        register_test(test_freeze_panes_vert);
        register_test(test_freeze_panes_both);
        register_test(test_lowest_column);
        register_test(test_lowest_column_or_props);
        register_test(test_highest_column);
        register_test(test_highest_column_or_props);
        register_test(test_lowest_row);
        register_test(test_lowest_row_or_props);
        register_test(test_highest_row);
        register_test(test_highest_row_or_props);
        register_test(test_const_iterators);
        register_test(test_const_reverse_iterators);
        register_test(test_column_major_iterators);
        register_test(test_reverse_column_major_iterators);
        register_test(test_const_column_major_iterators);
        register_test(test_const_reverse_column_major_iterators);
        register_test(test_header);
        register_test(test_footer);
        register_test(test_page_setup);
        register_test(test_unique_sheet_name);
        register_test(test_page_margins);
        register_test(test_garbage_collect);
        register_test(test_has_cell);
        register_test(test_get_range_by_string);
        register_test(test_operators);
        register_test(test_reserve);
        register_test(test_iterate);
        register_test(test_range_reference);
        register_test(test_get_point_pos);
        register_test(test_named_range_named_cell_reference);
        register_test(test_iteration_skip_empty);
        register_test(test_dimensions);
        register_test(test_view_properties_serialization);
        register_test(test_clear_cell);
        register_test(test_clear_row);
        register_test(test_set_title);
    }

    void test_new_worksheet()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        xlnt_assert(ws.workbook() == wb);
    }

    void test_cell()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        auto cell = ws.cell(xlnt::cell_reference(1, 1));
        xlnt_assert_equals(cell.reference(), "A1");
    }

    void test_invalid_cell()
    {
        xlnt_assert_throws(xlnt::cell_reference(xlnt::column_t((xlnt::column_t::index_t)0), 0),
            xlnt::invalid_cell_reference);
    }

    void test_worksheet_dimension()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();

        xlnt_assert_equals("A1:A1", ws.calculate_dimension());
        ws.cell("B12").value("AAA");
        xlnt_assert_equals("B12:B12", ws.calculate_dimension());
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

        xlnt_assert_equals(ws.calculate_dimension(), "A1:C9");
        xlnt_assert_equals(ws.rows(false)[row][column].reference(), coordinate);

        row = 8;
        column = 2;
        coordinate = "C9";

        xlnt_assert_equals(ws.rows(false)[row][column].reference(), coordinate);
    }

    void test_get_named_range()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        wb.create_named_range("test_range", ws, "C5");
        auto xlrange = ws.named_range("test_range");
        xlnt_assert_equals(1, xlrange.length());
        xlnt_assert_equals(1, xlrange[0].length());
        xlnt_assert_equals(5, xlrange[0][0].row());

        ws.create_named_range("test_range2", "C6");
        auto xlrange2 = ws.named_range("test_range2");
        xlnt_assert_equals(1, xlrange2.length());
        xlnt_assert_equals(1, xlrange2[0].length());
        xlnt_assert_equals(6, xlrange2[0][0].row());
    }

    void test_get_bad_named_range()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        xlnt_assert_throws(ws.named_range("bad_range"), xlnt::key_not_found);
    }

    void test_get_named_range_wrong_sheet()
    {
        xlnt::workbook wb;
        wb.create_sheet();
        wb.create_sheet();
        auto ws1 = wb[0];
        auto ws2 = wb[1];
        wb.create_named_range("wrong_sheet_range", ws1, "C5");
        xlnt_assert_throws(ws2.named_range("wrong_sheet_range"), xlnt::key_not_found);
    }

    void test_remove_named_range_bad()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        xlnt_assert_throws(ws.remove_named_range("bad_range"), std::runtime_error);
    }

    void test_cell_alternate_coordinates()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        auto cell = ws.cell(xlnt::cell_reference(4, 8));
        xlnt_assert_equals(cell.reference(), "D8");
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
        xlnt_assert_equals(c_range_coord, c_range_name);
        xlnt_assert(c_range_coord[0][0] == c_cell);
    }

    void test_rows()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();

        ws.cell("A1").value("first");
        ws.cell("C9").value("last");

        auto rows = ws.rows();

        xlnt_assert_equals(rows.length(), 9);

        auto first_row = rows[0];
        auto last_row = rows[8];

        xlnt_assert_equals(first_row[0].value<std::string>(), "first");
        xlnt_assert_equals(first_row[0].reference(), "A1");
        xlnt_assert_equals(last_row[2].value<std::string>(), "last");
    }

    void test_no_rows()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();

        xlnt_assert_equals(ws.rows().length(), 1);
        xlnt_assert_equals(ws.rows()[0].length(), 1);
    }

    void test_no_cols()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();

        xlnt_assert_equals(ws.columns().length(), 1);
        xlnt_assert_equals(ws.columns()[0].length(), 1);
    }

    void test_one_cell()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();

        auto cell = ws.cell("A1");

        xlnt_assert_equals(ws.columns().length(), 1);
        xlnt_assert_equals(ws.columns()[0].length(), 1);
        xlnt_assert_equals(ws.columns()[0][0], cell);
    }

    // void test_by_col() {}

    void test_cols()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();

        ws.cell("A1").value("first");
        ws.cell("C9").value("last");

        auto cols = ws.columns();

        xlnt_assert_equals(cols.length(), 3);

        xlnt_assert_equals(cols[0][0].value<std::string>(), "first");
        xlnt_assert_equals(cols[2][8].value<std::string>(), "last");
    }

    void test_auto_filter()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();

        ws.auto_filter(ws.range("a1:f1"));
        xlnt_assert_equals(ws.auto_filter(), "A1:F1");

        ws.clear_auto_filter();
        xlnt_assert(!ws.has_auto_filter());

        ws.auto_filter("c1:g9");
        xlnt_assert_equals(ws.auto_filter(), "C1:G9");
    }

    void test_getitem()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        xlnt::cell cell = ws[xlnt::cell_reference("A1")];
        xlnt_assert_equals(cell.reference().to_string(), "A1");
        xlnt_assert_equals(cell.data_type(), xlnt::cell::type::empty);
    }

    void test_setitem()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        ws[xlnt::cell_reference("A12")].value(5);
        xlnt_assert(ws[xlnt::cell_reference("A12")].value<int>() == 5);
    }

    void test_getslice()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        auto cell_range = ws.range("A1:B2");
        xlnt_assert_equals(cell_range[0][0], ws.cell("A1"));
        xlnt_assert_equals(cell_range[1][0], ws.cell("A2"));
        xlnt_assert_equals(cell_range[0][1], ws.cell("B1"));
        xlnt_assert_equals(cell_range[1][1], ws.cell("B2"));
    }

    void test_freeze()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();

        ws.freeze_panes(ws.cell("b2"));
        xlnt_assert_equals(ws.frozen_panes(), "B2");

        ws.unfreeze_panes();
        xlnt_assert(!ws.has_frozen_panes());

        ws.freeze_panes("c5");
        xlnt_assert_equals(ws.frozen_panes(), "C5");

        ws.freeze_panes(ws.cell("A1"));
        xlnt_assert(!ws.has_frozen_panes());
    }

    void test_merged_cells_lookup()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        ws.cell("A2").value("test");
        ws.merge_cells("A1:N50");
        auto all_merged = ws.merged_ranges();
        xlnt_assert_equals(all_merged.size(), 1);
        auto merged = ws.range(all_merged[0]);
        xlnt_assert(merged.contains("A1"));
        xlnt_assert(merged.contains("N50"));
        xlnt_assert(!merged.contains("A51"));
        xlnt_assert(!merged.contains("O1"));
    }

    void test_merged_cell_ranges()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        xlnt_assert_equals(ws.merged_ranges().size(), 0);
    }

    void test_merge_range_string()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        ws.cell("A1").value(1);
        ws.cell("D4").value(16);
        ws.merge_cells("A1:D4");
        std::vector<xlnt::range_reference> expected = {xlnt::range_reference("A1:D4")};
        xlnt_assert_equals(ws.merged_ranges(), expected);
        xlnt_assert(!ws.cell("D4").has_value());
    }

    void test_unmerge_bad()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();

        xlnt_assert_throws(ws.unmerge_cells("A1:D3"), std::runtime_error);
    }

    void test_unmerge_range_string()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        ws.merge_cells("A1:D4");
        xlnt_assert_equals(ws.merged_ranges().size(), 1);
        ws.unmerge_cells("A1:D4");
        xlnt_assert_equals(ws.merged_ranges().size(), 0);
    }

    void test_print_titles_old()
    {
        xlnt::workbook wb;

        auto ws = wb.active_sheet();
        ws.print_title_rows(3);
        xlnt_assert_equals(ws.print_titles(), "Sheet1!1:3");

        auto ws2 = wb.create_sheet();
        ws2.print_title_cols(4);
        xlnt_assert_equals(ws2.print_titles(), "Sheet2!A:D");
    }

    void test_print_titles_new()
    {
        xlnt::workbook wb;

        auto ws = wb.active_sheet();
        ws.print_title_rows(4);
        xlnt_assert_equals(ws.print_titles(), "Sheet1!1:4");

        auto ws2 = wb.create_sheet();
        ws2.print_title_cols("F");
        xlnt_assert_equals(ws2.print_titles(), "Sheet2!A:F");

        auto ws3 = wb.create_sheet();
        ws3.print_title_rows(2, 3);
        ws3.print_title_cols("C", "D");
        xlnt_assert_equals(ws3.print_titles(), "Sheet3!2:3,Sheet3!C:D");
    }

    void test_print_area()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        ws.print_area("A1:F5");
        xlnt_assert_equals(ws.print_area(), "$A$1:$F$5");
    }

    void test_freeze_panes_horiz()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        ws.freeze_panes("A4");

        auto view = ws.view();
        xlnt_assert_equals(view.selections().size(), 1);
        // pane is the corner of the worksheet that this selection extends to
        // active cell is the last selected cell in the selection
        // sqref is the last selected block in the selection
        xlnt_assert_equals(view.selections()[0].pane(), xlnt::pane_corner::bottom_left);
        xlnt_assert_equals(view.selections()[0].active_cell(), "A4");
        xlnt_assert_equals(view.selections()[0].sqref(), "A4");
        xlnt_assert_equals(view.pane().active_pane, xlnt::pane_corner::bottom_left);
        xlnt_assert_equals(view.pane().state, xlnt::pane_state::frozen);
        xlnt_assert_equals(view.pane().top_left_cell.get(), "A4");
        xlnt_assert_equals(view.pane().y_split, 3);
    }

    void test_freeze_panes_vert()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        ws.freeze_panes("D1");

        auto view = ws.view();
        xlnt_assert_equals(view.selections().size(), 1);
        // pane is the corner of the worksheet that this selection extends to
        // active cell is the last selected cell in the selection
        // sqref is the last selected block in the selection
        xlnt_assert_equals(view.selections()[0].pane(), xlnt::pane_corner::top_right);
        xlnt_assert_equals(view.selections()[0].active_cell(), "D1");
        xlnt_assert_equals(view.selections()[0].sqref(), "D1");
        xlnt_assert_equals(view.pane().active_pane, xlnt::pane_corner::top_right);
        xlnt_assert_equals(view.pane().state, xlnt::pane_state::frozen);
        xlnt_assert_equals(view.pane().top_left_cell.get(), "D1");
        xlnt_assert_equals(view.pane().x_split, 3);
    }

    void test_freeze_panes_both()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        ws.freeze_panes("D4");

        auto view = ws.view();
        xlnt_assert_equals(view.selections().size(), 3);
        // pane is the corner of the worksheet that this selection extends to
        // active cell is the last selected cell in the selection
        // sqref is the last selected block in the selection
        xlnt_assert_equals(view.selections()[0].pane(), xlnt::pane_corner::top_right);
        xlnt_assert_equals(view.selections()[0].active_cell(), "D1");
        xlnt_assert_equals(view.selections()[0].sqref(), "D1");
        xlnt_assert_equals(view.selections()[1].pane(), xlnt::pane_corner::bottom_left);
        xlnt_assert_equals(view.selections()[1].active_cell(), "A4");
        xlnt_assert_equals(view.selections()[1].sqref(), "A4");
        xlnt_assert_equals(view.selections()[2].pane(), xlnt::pane_corner::bottom_right);
        xlnt_assert_equals(view.selections()[2].active_cell(), "D4");
        xlnt_assert_equals(view.selections()[2].sqref(), "D4");
        xlnt_assert_equals(view.pane().active_pane, xlnt::pane_corner::bottom_right);
        xlnt_assert_equals(view.pane().state, xlnt::pane_state::frozen);
        xlnt_assert_equals(view.pane().top_left_cell.get(), "D4");
        xlnt_assert_equals(view.pane().x_split, 3);
        xlnt_assert_equals(view.pane().y_split, 3);
    }

    void test_lowest_column()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        xlnt_assert_equals(ws.lowest_column(), 1);
    }

    void test_lowest_column_or_props()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        ws.column_properties("J").width = 14.3;
        xlnt_assert_equals(ws.lowest_column_or_props(), "J");
    }

    void test_highest_column()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        ws[xlnt::cell_reference("F1")].value(10);
        ws[xlnt::cell_reference("F2")].value(32);
        ws[xlnt::cell_reference("F3")].formula("=F1+F2");
        ws[xlnt::cell_reference("A4")].formula("=A1+A2+A3");
        xlnt_assert_equals(ws.highest_column(), "F");
    }

    void test_highest_column_or_props()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        ws.column_properties("J").width = 14.3;
        xlnt_assert_equals(ws.highest_column_or_props(), "J");
    }

    void test_lowest_row()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        xlnt_assert_equals(ws.lowest_row(), 1);
    }

    void test_lowest_row_or_props()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        ws.row_properties(11).height = 14.3;
        xlnt_assert_equals(ws.lowest_row_or_props(), 11);
    }

    void test_highest_row()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        ws.cell("D4").value("D4");
        xlnt_assert_equals(ws.highest_row(), 4);
    }

    void test_highest_row_or_props()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        ws.row_properties(11).height = 14.3;
        xlnt_assert_equals(ws.highest_row_or_props(), 11);
    }

    void test_const_iterators()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();

        ws.cell("A1").value("A1");
        ws.cell("B1").value("B1");
        ws.cell("C1").value("C1");
        ws.cell("A2").value("A2");
        ws.cell("B2").value("B2");
        ws.cell("C2").value("C2");

        const xlnt::worksheet ws_const = ws;
        const auto rows = ws_const.rows();

        const auto first_row = rows.front();
        const auto first_cell = first_row.front();
        xlnt_assert_equals(first_cell.reference(), "A1");
        xlnt_assert_equals(first_cell.value<std::string>(), "A1");

        const auto last_row = rows.back();
        const auto last_cell = last_row.back();
        xlnt_assert_equals(last_cell.reference(), "C2");
        xlnt_assert_equals(last_cell.value<std::string>(), "C2");

        for (const auto row : rows)
        {
            for (const auto cell : row)
            {
                xlnt_assert_equals(cell.value<std::string>(), cell.reference().to_string());
            }
        }
    }

    void test_const_reverse_iterators()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();

        ws.cell("A1").value("A1");
        ws.cell("B1").value("B1");
        ws.cell("C1").value("C1");
        ws.cell("A2").value("A2");
        ws.cell("B2").value("B2");
        ws.cell("C2").value("C2");

        const xlnt::worksheet ws_const = ws;
        const auto rows = ws_const.rows();

        const auto first_row = *rows.rbegin();
        const auto first_cell = *first_row.rbegin();
        xlnt_assert_equals(first_cell.value<std::string>(), "C2");

        auto row_iter = rows.rend();
        row_iter--;
        const auto last_row = *row_iter;
        auto cell_iter = last_row.rend();
        cell_iter--;
        const auto last_cell = *cell_iter;
        xlnt_assert_equals(last_cell.value<std::string>(), "A1");

        for (auto ws_iter = rows.rbegin(); ws_iter != rows.rend(); ws_iter++)
        {
            const auto row = *ws_iter;

            for (auto row_iter = row.rbegin(); row_iter != row.rend(); row_iter++)
            {
                const auto cell = *row_iter;
                xlnt_assert_equals(cell.value<std::string>(), cell.reference().to_string());
            }
        }
    }

    void test_column_major_iterators()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();

        ws.cell("A1").value("A1");
        ws.cell("B1").value("B1");
        ws.cell("C1").value("C1");
        ws.cell("A2").value("A2");
        ws.cell("B2").value("B2");
        ws.cell("C2").value("C2");

        auto columns = ws.columns();

        auto first_column = *columns.begin();
        auto first_column_iter = first_column.begin();
        auto first_cell = *first_column_iter;
        xlnt_assert_equals(first_cell.value<std::string>(), "A1");
        first_column_iter++;
        auto second_cell = *first_column_iter;
        xlnt_assert_equals(second_cell.value<std::string>(), "A2");

        xlnt_assert_equals(first_cell, first_column.front());
        xlnt_assert_equals(second_cell, first_column.back());

        auto last_column = *(--columns.end());
        auto last_column_iter = last_column.end();
        last_column_iter--;
        auto last_cell = *last_column_iter;
        xlnt_assert_equals(last_cell.value<std::string>(), "C2");
        last_column_iter--;
        auto penultimate_cell = *last_column_iter;
        xlnt_assert_equals(penultimate_cell.value<std::string>(), "C1");

        for (auto column : columns)
        {
            for (auto cell : column)
            {
                xlnt_assert_equals(cell.value<std::string>(), cell.reference().to_string());
            }
        }
    }

    void test_reverse_column_major_iterators()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();

        ws.cell("A1").value("A1");
        ws.cell("B1").value("B1");
        ws.cell("C1").value("C1");
        ws.cell("A2").value("A2");
        ws.cell("B2").value("B2");
        ws.cell("C2").value("C2");

        auto columns = ws.columns();

        auto column_iter = columns.rbegin();
        auto first_column = *column_iter;

        auto first_column_iter = first_column.rbegin();
        auto first_cell = *first_column_iter;
        xlnt_assert_equals(first_cell.value<std::string>(), "C2");
        first_column_iter++;
        auto second_cell = *first_column_iter;
        xlnt_assert_equals(second_cell.value<std::string>(), "C1");

        auto last_column = *(--columns.rend());
        auto last_column_iter = last_column.rend();
        last_column_iter--;
        auto last_cell = *last_column_iter;
        xlnt_assert_equals(last_cell.value<std::string>(), "A1");
        last_column_iter--;
        auto penultimate_cell = *last_column_iter;
        xlnt_assert_equals(penultimate_cell.value<std::string>(), "A2");

        for (auto column_iter = columns.rbegin(); column_iter != columns.rend(); ++column_iter)
        {
            auto column = *column_iter;

            for (auto cell_iter = column.rbegin(); cell_iter != column.rend(); ++cell_iter)
            {
                auto cell = *cell_iter;

                xlnt_assert_equals(cell.value<std::string>(), cell.reference().to_string());
            }
        }
    }

    void test_const_column_major_iterators()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();

        ws.cell("A1").value("A1");
        ws.cell("B1").value("B1");
        ws.cell("C1").value("C1");
        ws.cell("A2").value("A2");
        ws.cell("B2").value("B2");
        ws.cell("C2").value("C2");

        const xlnt::worksheet ws_const = ws;
        const auto columns = ws_const.columns();

        const auto first_column = *columns.begin();
        auto first_column_iter = first_column.begin();
        const auto first_cell = *first_column_iter;
        xlnt_assert_equals(first_cell.value<std::string>(), "A1");
        first_column_iter++;
        const auto second_cell = *first_column_iter;
        xlnt_assert_equals(second_cell.value<std::string>(), "A2");

        const auto last_column = *(--columns.end());
        auto last_column_iter = last_column.end();
        last_column_iter--;
        const auto last_cell = *last_column_iter;
        xlnt_assert_equals(last_cell.value<std::string>(), "C2");
        last_column_iter--;
        const auto penultimate_cell = *last_column_iter;
        xlnt_assert_equals(penultimate_cell.value<std::string>(), "C1");

        for (const auto column : columns)
        {
            for (const auto cell : column)
            {
                xlnt_assert_equals(cell.value<std::string>(), cell.reference().to_string());
            }
        }
    }

    void test_const_reverse_column_major_iterators()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();

        ws.cell("A1").value("A1");
        ws.cell("B1").value("B1");
        ws.cell("C1").value("C1");
        ws.cell("A2").value("A2");
        ws.cell("B2").value("B2");
        ws.cell("C2").value("C2");

        const xlnt::worksheet ws_const = ws;
        const auto columns = ws_const.columns();

        const auto first_column = *columns.crbegin();
        auto first_column_iter = first_column.crbegin();
        const auto first_cell = *first_column_iter;
        xlnt_assert_equals(first_cell.value<std::string>(), "C2");
        first_column_iter++;
        const auto second_cell = *first_column_iter;
        xlnt_assert_equals(second_cell.value<std::string>(), "C1");

        const auto last_column = *(--columns.crend());
        auto last_column_iter = last_column.crend();
        last_column_iter--;
        const auto last_cell = *last_column_iter;
        xlnt_assert_equals(last_cell.value<std::string>(), "A1");
        last_column_iter--;
        const auto penultimate_cell = *last_column_iter;
        xlnt_assert_equals(penultimate_cell.value<std::string>(), "A2");

        for (auto column_iter = columns.crbegin(); column_iter != columns.crend(); ++column_iter)
        {
            const auto column = *column_iter;

            for (auto cell_iter = column.crbegin(); cell_iter != column.crend(); ++cell_iter)
            {
                const auto cell = *cell_iter;

                xlnt_assert_equals(cell.value<std::string>(), cell.reference().to_string());
            }
        }
    }

    void test_header()
    {
        xlnt::header_footer hf;
        using hf_loc = xlnt::header_footer::location;

        for (auto location : {hf_loc::left, hf_loc::center, hf_loc::right})
        {
            xlnt_assert(!hf.has_header(location));
            xlnt_assert(!hf.has_odd_even_header(location));
            xlnt_assert(!hf.has_first_page_header(location));

            hf.header(location, "abc");

            xlnt_assert(hf.has_header(location));
            xlnt_assert(!hf.has_odd_even_header(location));
            xlnt_assert(!hf.has_first_page_header(location));

            xlnt_assert_equals(hf.header(location), "abc");

            hf.clear_header(location);

            xlnt_assert(!hf.has_header(location));
        }
    }

    void test_footer()
    {
        xlnt::header_footer hf;
        using hf_loc = xlnt::header_footer::location;

        for (auto location : {hf_loc::left, hf_loc::center, hf_loc::right})
        {
            xlnt_assert(!hf.has_footer(location));
            xlnt_assert(!hf.has_odd_even_footer(location));
            xlnt_assert(!hf.has_first_page_footer(location));

            hf.footer(location, "abc");

            xlnt_assert(hf.has_footer(location));
            xlnt_assert(!hf.has_odd_even_footer(location));
            xlnt_assert(!hf.has_first_page_footer(location));

            xlnt_assert_equals(hf.footer(location), "abc");

            hf.clear_footer(location);

            xlnt_assert(!hf.has_footer(location));
        }
    }

    void test_page_setup()
    {
        xlnt::page_setup setup;
        setup.page_break(xlnt::page_break::column);
        xlnt_assert_equals(setup.page_break(), xlnt::page_break::column);
        setup.scale(1.23);
        xlnt_assert_equals(setup.scale(), 1.23);
    }

    void test_unique_sheet_name()
    {
        xlnt::workbook wb;

        auto first_created = wb.create_sheet();
        auto second_created = wb.create_sheet();

        xlnt_assert_differs(first_created.title(), second_created.title());
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

        ws.page_margins(margins);

        xlnt_assert(ws.has_page_margins());
        xlnt_assert_equals(ws.page_margins().top(), 0);
        xlnt_assert_equals(ws.page_margins().bottom(), 1);
        xlnt_assert_equals(ws.page_margins().header(), 2);
        xlnt_assert_equals(ws.page_margins().footer(), 3);
        xlnt_assert_equals(ws.page_margins().left(), 4);
        xlnt_assert_equals(ws.page_margins().right(), 5);
    }

    void test_garbage_collect()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();

        auto dimensions = ws.calculate_dimension();
        xlnt_assert_equals(dimensions, xlnt::range_reference("A1", "A1"));

        ws.cell("B2").value("text");
        ws.garbage_collect();

        dimensions = ws.calculate_dimension();
        xlnt_assert_equals(dimensions, xlnt::range_reference("B2", "B2"));
    }

    void test_has_cell()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();

        ws.cell("A3").value("test");

        xlnt_assert(!ws.has_cell("A2"));
        xlnt_assert(ws.has_cell("A3"));
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
        xlnt_assert_equals((*row_iter).value<double>(), 3.14);
        xlnt_assert_equals((*row_iter).reference(), "A2");
        xlnt_assert_equals((*row_iter), row.front());

        row_iter++;
        xlnt_assert_equals((*row_iter).value<std::string>(), "text");
        xlnt_assert_equals((*row_iter).reference(), "B2");
        xlnt_assert_equals((*row_iter), row.back());

        range_iter++;
        row = *range_iter;
        row_iter = row.begin();
        xlnt_assert_equals((*row_iter).value<bool>(), true);
        xlnt_assert_equals((*row_iter).reference(), "A3");

        range_iter = range.end();
        range_iter--;
        row = *range_iter;
        row_iter = row.end();
        row_iter--;
        xlnt_assert_equals((*row_iter).value<bool>(), false);
    }

    void test_operators()
    {
        xlnt::workbook wb;

        wb.create_sheet();
        wb.create_sheet();

        auto ws1 = wb[1];
        auto ws2 = wb[2];

        xlnt_assert_differs(ws1, ws2);

        ws1[xlnt::cell_reference("A2")].value(true);

        xlnt_assert_equals(ws1[xlnt::cell_reference("A2")].value<bool>(), true);
        xlnt_assert_equals((*(*ws1.range("A2:A2").begin()).begin()).value<bool>(), true);

        ws1.create_named_range("rangey", "A2:A2");
        xlnt_assert_equals(ws1.range("rangey"), ws1.range("A2:A2"));
        xlnt_assert_equals(ws1.range("A2:A2"), ws1.range("A2:A2"));
        xlnt_assert(ws1.range("rangey") != ws1.range("A2:A3"));

        xlnt_assert_equals(ws1.range("rangey").cell("A1"), ws1.cell("A2"));
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
                    xlnt_assert_equals(cell.reference().to_string(), cell.value<std::string>());
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
                    xlnt_assert_equals(cell.reference().to_string(), cell.value<std::string>());
                }
            }
        }

        const auto const_range = ws_const.range("B3:C7");
        auto const_range_iter = const_range.cbegin();
        const_range_iter++;
        const_range_iter--;
        xlnt_assert_equals(const_range_iter, const_range.begin());
    }

    void test_range_reference()
    {
        xlnt::range_reference ref1("A1:A1");
        xlnt_assert(ref1.is_single_cell());
        xlnt::range_reference ref2("A1:B2");

        xlnt_assert(!ref2.is_single_cell());
        xlnt_assert(ref1 == xlnt::range_reference("A1:A1"));
        xlnt_assert(ref1 != ref2);
        xlnt_assert(ref1 == "A1:A1");
        xlnt_assert(ref1 == std::string("A1:A1"));
        xlnt_assert(std::string("A1:A1") == ref1);
        xlnt_assert("A1:A1" == ref1);
        xlnt_assert(ref1 != "A1:B2");
        xlnt_assert(ref1 != std::string("A1:B2"));
        xlnt_assert(std::string("A1:B2") != ref1);
        xlnt_assert("A1:B2" != ref1);
    }

    void test_get_point_pos()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();

        xlnt_assert_equals(ws.point_pos(0, 0), "A1");
    }

    void test_named_range_named_cell_reference()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        xlnt_assert_throws(ws.create_named_range("A1", "A2"), xlnt::invalid_parameter);
        xlnt_assert_throws(ws.create_named_range("XFD1048576", "A2"), xlnt::invalid_parameter);
        xlnt_assert_throws_nothing(ws.create_named_range("XFE1048576", "A2"));
        xlnt_assert_throws_nothing(ws.create_named_range("XFD1048577", "A2"));
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

            xlnt_assert_equals(cells.size(), 2);
            xlnt_assert_equals(cells[0].value<std::string>(), "A1");
            xlnt_assert_equals(cells[1].value<std::string>(), "F6");
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

            xlnt_assert_equals(cells.size(), 2);
            xlnt_assert_equals(cells[0].value<std::string>(), "A1");
            xlnt_assert_equals(cells[1].value<std::string>(), "F6");
        }
    }

    void test_dimensions()
    {
        xlnt::workbook workbook;
        workbook.load(path_helper::test_file("4_every_style.xlsx"));

        auto active_sheet = workbook.active_sheet();
        auto sheet_range = active_sheet.calculate_dimension();

        xlnt_assert(!sheet_range.is_single_cell());
        xlnt_assert_equals(sheet_range.width(), 4);
        xlnt_assert_equals(sheet_range.height(), 35);
    }

    void test_view_properties_serialization()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();

        ws.cell("A1").value("A1");
        ws.cell("A2").value("A2");
        ws.cell("B1").value("B1");
        ws.cell("B2").value("B2");

        xlnt_assert(!ws.has_active_cell());

        ws.active_cell("B1");

        xlnt_assert(ws.has_active_cell());
        xlnt_assert_equals(ws.active_cell(), "B1");

        wb.save("temp.xlsx");

        xlnt::workbook wb2;
        wb2.load("temp.xlsx");
        auto ws2 = wb2.active_sheet();

        xlnt_assert(ws2.has_active_cell());
        xlnt_assert_equals(ws2.active_cell(), "B1");
    }

    void test_clear_cell()
    {
        xlnt::workbook wb;
        wb.load(path_helper::test_file("4_every_style.xlsx"));

        auto ws = wb.active_sheet();
        auto height = ws.calculate_dimension().height();
        auto last_row = ws.highest_row();

        xlnt_assert(ws.has_cell(xlnt::cell_reference(1, last_row)));

        ws.clear_cell(xlnt::cell_reference(1, last_row));

        xlnt_assert(!ws.has_cell(xlnt::cell_reference(1, last_row)));
        xlnt_assert_equals(ws.highest_row(), last_row);

        wb.save("temp.xlsx");

        xlnt::workbook wb2;
        wb2.load("temp.xlsx");
        auto ws2 = wb2.active_sheet();

        xlnt_assert_equals(ws2.calculate_dimension().height(), height);
        xlnt_assert(!ws2.has_cell(xlnt::cell_reference(1, last_row)));
    }

    void test_clear_row()
    {
        xlnt::workbook wb;
        wb.load(path_helper::test_file("4_every_style.xlsx"));

        auto ws = wb.active_sheet();
        auto height = ws.calculate_dimension().height();
        auto last_row = ws.highest_row();

        xlnt_assert(ws.has_cell(xlnt::cell_reference(1, last_row)));

        ws.clear_row(last_row);

        xlnt_assert(!ws.has_cell(xlnt::cell_reference(1, last_row)));
        xlnt_assert_equals(ws.highest_row(), last_row - 1);

        wb.save("temp.xlsx");

        xlnt::workbook wb2;
        wb2.load("temp.xlsx");
        auto ws2 = wb2.active_sheet();

        xlnt_assert_equals(ws2.calculate_dimension().height(), height - 1);
        xlnt_assert(!ws2.has_cell(xlnt::cell_reference(1, last_row)));
    }

    void test_set_title()
    {
        xlnt::workbook wb;
        auto ws1 = wb.active_sheet();
        // empty titles are invalid
        xlnt_assert_throws(ws1.title(""), xlnt::invalid_sheet_title);
        // titles longer than 31 chars are invalid
        std::string test_long_title(32, 'a');
        xlnt_assert(test_long_title.size() > 31);
        xlnt_assert_throws(ws1.title(test_long_title), xlnt::invalid_sheet_title);
        // titles containing any of the following characters are invalid
        std::string invalid_chars = "*:/\\?[]";
        for (char &c : invalid_chars)
        {
            std::string invalid_char = std::string("Sheet") + c;
            xlnt_assert_throws(ws1.title(invalid_char), xlnt::invalid_sheet_title);
        }
        // duplicate names are invalid
        auto ws2 = wb.create_sheet();
        xlnt_assert_throws(ws2.title(ws1.title()), xlnt::invalid_sheet_title);
        xlnt_assert_throws(ws1.title(ws2.title()), xlnt::invalid_sheet_title);
        // naming as self is valid and is ignored
        auto ws1_title = ws1.title();
        auto ws2_title = ws2.title();
        xlnt_assert_throws_nothing(ws1.title(ws1.title()));
        xlnt_assert_throws_nothing(ws2.title(ws2.title()));
        xlnt_assert(ws1_title == ws1.title());
        xlnt_assert(ws2_title == ws2.title());
    }
};
static worksheet_test_suite x;