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


#include <helpers/test_suite.hpp>
#include <xlnt/cell/cell.hpp>
#include <xlnt/styles/font.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/worksheet/header_footer.hpp>
#include <xlnt/worksheet/range.hpp>
#include <xlnt/worksheet/worksheet.hpp>

class range_test_suite : public test_suite
{
public:
    range_test_suite()
    {
        register_test(test_construction);
        register_test(test_batch_formatting);
        register_test(test_clear_cells);
    }

    void test_construction()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();

        xlnt::range range_1(ws, xlnt::range_reference("A1:D10"));
        xlnt_assert_equals(range_1.target_worksheet(), ws);
        xlnt_assert_equals(1, range_1.front()[0].row()); // NOTE: querying row/column here desperately needs some shortcuts
        xlnt_assert_equals(xlnt::column_t("D"), range_1.front().back().column());
        xlnt_assert_equals(10, range_1.back()[0].row());
        xlnt_assert_equals(xlnt::column_t("D"), range_1.back().back().column());
        // assert default parameters in ctor
        xlnt::range range_2(ws, xlnt::range_reference("A1:D10"), xlnt::major_order::row, false);
        xlnt_assert_equals(range_1, range_2);
        // assert copy
        xlnt::range range_3(range_2);
        xlnt_assert_equals(range_1, range_3);

        // column order
        xlnt::range range_4(ws, xlnt::range_reference("A1:D10"), xlnt::major_order::column);
        xlnt_assert_equals(xlnt::column_t("A"), range_4.front()[0].column()); // NOTE: querying row/column here desperately needs some shortcuts
        xlnt_assert_equals(10, range_4.front().back().row());
        xlnt_assert_equals(xlnt::column_t("D"), range_4.back()[0].column());
        xlnt_assert_equals(10, range_4.back().back().row());
        // assignment
        range_3 = range_4;
        xlnt_assert_equals(range_3, range_4);
    }

    void test_batch_formatting()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();

        for (auto row = 1; row <= 10; ++row)
        {
            for (auto column = 1; column <= 10; ++column)
            {
                auto ref = xlnt::cell_reference(column, row);
                ws[ref].value(ref.to_string());
            }
        }

        ws.range("A1:A10").font(xlnt::font().name("Arial"));
        ws.range("A1:J1").font(xlnt::font().bold(true));

        xlnt_assert_equals(ws.cell("A1").font().name(), "Calibri");
        xlnt_assert(ws.cell("A1").font().bold());

        xlnt_assert_equals(ws.cell("A2").font().name(), "Arial");
        xlnt_assert(!ws.cell("A2").font().bold());

        xlnt_assert_equals(ws.cell("B1").font().name(), "Calibri");
        xlnt_assert(ws.cell("B1").font().bold());

        xlnt_assert(!ws.cell("B2").has_format());
    }
    
    void test_clear_cells()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        ws.cell("A1").value("A1");
        ws.cell("A3").value("A3");
        ws.cell("C1").value("C1");
        ws.cell("B2").value("B2");
        ws.cell("C3").value("C3");
        xlnt_assert_equals(ws.calculate_dimension(), xlnt::range_reference(1, 1, 3, 3));
        auto range = ws.range("B1:C3");
        range.clear_cells();
        xlnt_assert_equals(ws.calculate_dimension(), xlnt::range_reference(1, 1, 1, 3));
    }
};
static range_test_suite x;