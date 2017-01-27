#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/xlnt.hpp>

class test_fill : public CxxTest::TestSuite
{
public:
    void test_properties()
    {
        xlnt::fill fill;
        
        TS_ASSERT_EQUALS(fill.type(), xlnt::fill_type::pattern);
        fill = xlnt::fill(xlnt::gradient_fill());
        TS_ASSERT_EQUALS(fill.type(), xlnt::fill_type::gradient);
        TS_ASSERT_EQUALS(fill.gradient_fill().type(), xlnt::gradient_fill_type::linear);

        TS_ASSERT_EQUALS(fill.gradient_fill().left(), 0);
        TS_ASSERT_EQUALS(fill.gradient_fill().right(), 0);
        TS_ASSERT_EQUALS(fill.gradient_fill().top(), 0);
        TS_ASSERT_EQUALS(fill.gradient_fill().bottom(), 0);

        TS_ASSERT_THROWS(fill.pattern_fill(), xlnt::invalid_attribute);

        TS_ASSERT_EQUALS(fill.gradient_fill().degree(), 0);
		fill = fill.gradient_fill().degree(1);
        TS_ASSERT_EQUALS(fill.gradient_fill().degree(), 1);

        fill = xlnt::pattern_fill().type(xlnt::pattern_fill_type::solid);

        fill = fill.pattern_fill().foreground(xlnt::color::black());
        TS_ASSERT(fill.pattern_fill().foreground().is_set());
        TS_ASSERT_EQUALS(fill.pattern_fill().foreground().get().rgb().hex_string(),
            xlnt::color::black().rgb().hex_string());

        fill = fill.pattern_fill().background(xlnt::color::green());
        TS_ASSERT(fill.pattern_fill().background().is_set());
        TS_ASSERT_EQUALS(fill.pattern_fill().background().get().rgb().hex_string(),
            xlnt::color::green().rgb().hex_string());
        
        const auto &const_fill = fill;
        TS_ASSERT(const_fill.pattern_fill().foreground().is_set());
        TS_ASSERT(const_fill.pattern_fill().background().is_set());
    }
    
    void test_comparison()
    {
        xlnt::fill pattern_fill = xlnt::pattern_fill().type(xlnt::pattern_fill_type::solid);
        xlnt::fill gradient_fill_linear = xlnt::gradient_fill().type(xlnt::gradient_fill_type::linear);
        xlnt::fill gradient_fill_path = xlnt::gradient_fill().type(xlnt::gradient_fill_type::path);

        TS_ASSERT_DIFFERS(pattern_fill, gradient_fill_linear);
        TS_ASSERT_DIFFERS(gradient_fill_linear, gradient_fill_path);
        TS_ASSERT_DIFFERS(gradient_fill_path, pattern_fill);
    }

    void test_two_fills()
    {
        xlnt::workbook wb;
        auto ws1 = wb.active_sheet();

        auto cell1 = ws1.cell("A10");
        auto cell2 = ws1.cell("A11");

        cell1.fill(xlnt::fill::solid(xlnt::color::yellow()));
        cell1.value("Fill Yellow");

        cell2.fill(xlnt::fill::solid(xlnt::color::green()));
        cell2.value(xlnt::date(2010, 7, 13));

        TS_ASSERT_EQUALS(cell1.fill(), xlnt::fill::solid(xlnt::color::yellow()));
        TS_ASSERT_EQUALS(cell2.fill(), xlnt::fill::solid(xlnt::color::green()));
    }
};
