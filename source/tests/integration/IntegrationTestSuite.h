#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/workbook.h>
#include <xlnt/worksheet.h>

class IntegrationTestSuite : public CxxTest::TestSuite
{
public:
    IntegrationTestSuite()
    {

    }

    void test_1()
    {
        auto wb = xlnt::workbook();
        auto ws = wb.get_active();
        auto ws1 = wb.create_sheet();
        auto ws2 = wb.create_sheet(0);
        ws.set_title("New Title");
        auto ws3 = wb["New Title"];
        auto ws4 = wb.get_sheet_by_name("New Title");
        TS_ASSERT_EQUALS(ws, ws3);
        TS_ASSERT_EQUALS(ws, ws4);
        TS_ASSERT_EQUALS(ws3, ws4);
        auto sheet_names = wb.get_sheet_names();
        for(auto sheet : wb)
        {
            std::cout << sheet.get_title() << std::endl;
        }

        auto cell_range = ws["A1:C2"];

        auto c = ws["A4"];
        ws["A4"].value() = 4;
        auto d = ws.cell(4, 2);
        c.value() = "hello, world";
        std::cout << (std::string)c.value() << std::endl;
        d.value() = 3.14;
        std::cout << (double)d.value() << std::endl;

        wb.save("balances.xlsx");
    }
};
