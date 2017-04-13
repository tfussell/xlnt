#pragma once

#include <iostream>
#include <helpers/test_suite.hpp>

#include <xlnt/worksheet/header_footer.hpp>
#include <xlnt/worksheet/worksheet.hpp>
#include <xlnt/workbook/workbook.hpp>

class test_range : public test_suite
{
public:
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

        assert_equals(ws.cell("A1").font().name(), "Calibri");
        assert(ws.cell("A1").font().bold());

        assert_equals(ws.cell("A2").font().name(), "Arial");
        assert(!ws.cell("A2").font().bold());

        assert_equals(ws.cell("B1").font().name(), "Calibri");
        assert(ws.cell("B1").font().bold());

        assert(!ws.cell("B2").has_format());
    }
};
