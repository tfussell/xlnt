#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include "../xlnt.h"

class IterTestSuite : public CxxTest::TestSuite
{
public:
    IterTestSuite()
    {

    }

    void test_get_dimensions()
    {
        auto expected = {"A1:G5", "D1:K30", "D2:D2", "A1:C1"};

        wb = _open_wb();

        for(i, sheetn : enumerate(wb.get_sheet_names()))
        {
            ws = wb.get_sheet_by_name(name = sheetn);
            TS_ASSERT_EQUALS(ws._dimensions, expected[i]);
        }
    }

    void test_read_fast_integrated()
    {
        std::string sheet_name = "Sheet1 - Text";

        std::vector<std::vector<char *>> expected = {{"This is cell A1 in Sheet 1", nullptr, nullptr, nullptr, nullptr, nullptr, nullptr},
        {nullptr, nullptr, nullptr, nullptr, nullptr, nullptr, nullptr},
        {nullptr, nullptr, nullptr, nullptr, nullptr, nullptr, nullptr},
        {nullptr, nullptr, nullptr, nullptr, nullptr, nullptr, nullptr},
        {nullptr, nullptr, nullptr, nullptr, nullptr, nullptr, "This is cell G5"}};

        wb = load_workbook(filename = workbook_name, use_iterators = True);
        ws = wb.get_sheet_by_name(name = sheet_name);

        for(row, expected_row : zip(ws.iter_rows(), expected)
        {
            row_values = [x.internal_value for x in row];
            TS_ASSERT_EQUALS(row_values, expected_row);
        }
    }

    void test_get_boundaries_range()
    {
        TS_ASSERT_EQUALS(get_range_boundaries("C1:C4"), (3, 1, 3, 4));
    }

    void test_get_boundaries_one()
    {
        TS_ASSERT_EQUALS(get_range_boundaries("C1"), (3, 1, 4, 1));
    }

    void test_read_single_cell_range()
    {
        wb = load_workbook(filename = workbook_name, use_iterators = True);
        ws = wb.get_sheet_by_name(name = sheet_name);

        TS_ASSERT_EQUALS("This is cell A1 in Sheet 1", list(ws.iter_rows("A1"))[0][0].internal_value);
    }

    void test_read_fast_integrated2()
    {
        sheet_name = "Sheet2 - Numbers";

        expected = [[x + 1] for x in range(30)];

        query_range = "D1:E30";

        wb = load_workbook(filename = workbook_name, use_iterators = True);
        ws = wb.get_sheet_by_name(name = sheet_name);

        for(row, expected_row : zip(ws.iter_rows(query_range), expected))
        {
            row_values = [x.internal_value for x in row];
            TS_ASSERT_EQUALS(row_values, expected_row);
        }
    }

    void test_read_single_cell_date()
    {
        sheet_name = "Sheet4 - Dates";

        wb = load_workbook(filename = workbook_name, use_iterators = True);
        ws = wb.get_sheet_by_name(name = sheet_name);

        TS_ASSERT_EQUALS(datetime.datetime(1973, 5, 20), list(ws.iter_rows("A1"))[0][0].internal_value);
        TS_ASSERT_EQUALS(datetime.datetime(1973, 5, 20, 9, 15, 2), list(ws.iter_rows("C1"))[0][0].internal_value);
    }
};
