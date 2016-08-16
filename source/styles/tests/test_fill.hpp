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
        TS_ASSERT(fill.pattern_fill().foreground());
        TS_ASSERT_EQUALS((*fill.pattern_fill().foreground()).get_rgb().get_hex_string(), xlnt::color::black().get_rgb().get_hex_string());

        fill = fill.pattern_fill().background(xlnt::color::green());
        TS_ASSERT(fill.pattern_fill().background());
        TS_ASSERT_EQUALS((*fill.pattern_fill().background()).get_rgb().get_hex_string(), xlnt::color::green().get_rgb().get_hex_string());
        
        const auto &const_fill = fill;
        TS_ASSERT(const_fill.pattern_fill().foreground());
        TS_ASSERT(const_fill.pattern_fill().background());
    }
    
    void test_hash()
    {
        xlnt::fill pattern_fill = xlnt::pattern_fill().type(xlnt::pattern_fill_type::solid);
        xlnt::fill gradient_fill_linear = xlnt::gradient_fill().type(xlnt::gradient_fill_type::linear);
        xlnt::fill gradient_fill_path = xlnt::gradient_fill().type(xlnt::gradient_fill_type::path);

        TS_ASSERT_DIFFERS(pattern_fill.hash(), gradient_fill_linear.hash());
        TS_ASSERT_DIFFERS(gradient_fill_linear.hash(), gradient_fill_path.hash());
        TS_ASSERT_DIFFERS(gradient_fill_path.hash(), pattern_fill.hash());
    }
};
