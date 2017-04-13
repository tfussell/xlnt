#pragma once

#include <iostream>
#include <helpers/test_suite.hpp>

#include <xlnt/xlnt.hpp>

class test_fill : public test_suite
{
public:
    void test_properties()
    {
        xlnt::fill fill;
        
        assert_equals(fill.type(), xlnt::fill_type::pattern);
        fill = xlnt::fill(xlnt::gradient_fill());
        assert_equals(fill.type(), xlnt::fill_type::gradient);
        assert_equals(fill.gradient_fill().type(), xlnt::gradient_fill_type::linear);

        assert_equals(fill.gradient_fill().left(), 0);
        assert_equals(fill.gradient_fill().right(), 0);
        assert_equals(fill.gradient_fill().top(), 0);
        assert_equals(fill.gradient_fill().bottom(), 0);

        assert_throws(fill.pattern_fill(), xlnt::invalid_attribute);

        assert_equals(fill.gradient_fill().degree(), 0);
		fill = fill.gradient_fill().degree(1);
        assert_equals(fill.gradient_fill().degree(), 1);

        fill = xlnt::pattern_fill().type(xlnt::pattern_fill_type::solid);

        fill = fill.pattern_fill().foreground(xlnt::color::black());
        assert(fill.pattern_fill().foreground().is_set());
        assert_equals(fill.pattern_fill().foreground().get().rgb().hex_string(),
            xlnt::color::black().rgb().hex_string());

        fill = fill.pattern_fill().background(xlnt::color::green());
        assert(fill.pattern_fill().background().is_set());
        assert_equals(fill.pattern_fill().background().get().rgb().hex_string(),
            xlnt::color::green().rgb().hex_string());
        
        const auto &const_fill = fill;
        assert(const_fill.pattern_fill().foreground().is_set());
        assert(const_fill.pattern_fill().background().is_set());
    }
    
    void test_comparison()
    {
        xlnt::fill pattern_fill = xlnt::pattern_fill().type(xlnt::pattern_fill_type::solid);
        xlnt::fill gradient_fill_linear = xlnt::gradient_fill().type(xlnt::gradient_fill_type::linear);
        xlnt::fill gradient_fill_path = xlnt::gradient_fill().type(xlnt::gradient_fill_type::path);

        assert_differs(pattern_fill, gradient_fill_linear);
        assert_differs(gradient_fill_linear, gradient_fill_path);
        assert_differs(gradient_fill_path, pattern_fill);
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

        assert_equals(cell1.fill(), xlnt::fill::solid(xlnt::color::yellow()));
        assert_equals(cell2.fill(), xlnt::fill::solid(xlnt::color::green()));
    }
};
