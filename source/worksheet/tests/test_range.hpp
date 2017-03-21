#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/worksheet/header_footer.hpp>
#include <xlnt/worksheet/worksheet.hpp>
#include <xlnt/workbook/workbook.hpp>

class test_range : public CxxTest::TestSuite
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

        TS_ASSERT_EQUALS(ws.cell("A1").font().name(), "Calibri");
        TS_ASSERT(ws.cell("A1").font().bold());

        TS_ASSERT_EQUALS(ws.cell("A2").font().name(), "Arial");
        TS_ASSERT(!ws.cell("A2").font().bold());

        TS_ASSERT_EQUALS(ws.cell("B1").font().name(), "Calibri");
        TS_ASSERT(ws.cell("B1").font().bold());

        TS_ASSERT(!ws.cell("B2").has_format());
    }
};
