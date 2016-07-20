#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <detail/style_serializer.hpp>
#include <detail/stylesheet.hpp>
#include <detail/workbook_impl.hpp>
#include <helpers/path_helper.hpp>

class test_style_reader : public CxxTest::TestSuite
{
public:
    void test_complex_formatting()
    {
        xlnt::workbook wb;
        wb.load(path_helper::get_data_directory("/reader/formatting.xlsx"));
        
        // border_style
        TS_ASSERT_EQUALS(wb.get_active_sheet().get_cell("E30").get_border().get_top()->get_color(), xlnt::color(xlnt::color::type::indexed, 10));
        TS_ASSERT_EQUALS(wb.get_active_sheet().get_cell("E30").get_border().get_top()->get_border_style(), xlnt::border_style::thin);
        
        // underline_style
        TS_ASSERT_EQUALS(wb.get_active_sheet().get_cell("E30").get_font().get_underline(), xlnt::font::underline_style::none);
        TS_ASSERT_EQUALS(wb.get_active_sheet().get_cell("F30").get_font().get_underline(), xlnt::font::underline_style::single);
    }
};
