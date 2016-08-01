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
        
        TS_ASSERT_EQUALS(fill.get_type(), xlnt::fill::type::pattern);
        fill = xlnt::fill::gradient(xlnt::gradient_fill::type::linear);
        TS_ASSERT_EQUALS(fill.get_type(), xlnt::fill::type::gradient);
        TS_ASSERT_EQUALS(fill.get_gradient_fill().get_type(), xlnt::gradient_fill::type::linear);

        TS_ASSERT_EQUALS(fill.get_gradient_fill().get_gradient_left(), 0);
        TS_ASSERT_EQUALS(fill.get_gradient_fill().get_gradient_right(), 0);
        TS_ASSERT_EQUALS(fill.get_gradient_fill().get_gradient_top(), 0);
        TS_ASSERT_EQUALS(fill.get_gradient_fill().get_gradient_bottom(), 0);

        TS_ASSERT_THROWS(fill.get_pattern_fill(), std::runtime_error);

        TS_ASSERT_EQUALS(fill.get_gradient_fill().get_degree(), 0);
        fill.get_gradient_fill().set_degree(1);
        TS_ASSERT_EQUALS(fill.get_gradient_fill().get_degree(), 1);

        fill = xlnt::fill::pattern(xlnt::pattern_fill::type::solid);

        fill.get_pattern_fill().set_foreground_color(xlnt::color::black());
        TS_ASSERT(fill.get_pattern_fill().get_foreground_color());
        TS_ASSERT_EQUALS((*fill.get_pattern_fill().get_foreground_color()).get_rgb().get_hex_string(), xlnt::color::black().get_rgb().get_hex_string());

        fill.get_pattern_fill().set_background_color(xlnt::color::green());
        TS_ASSERT(fill.get_pattern_fill().get_background_color());
        TS_ASSERT_EQUALS((*fill.get_pattern_fill().get_background_color()).get_rgb().get_hex_string(), xlnt::color::green().get_rgb().get_hex_string());
        
        const auto &const_fill = fill;
        TS_ASSERT(const_fill.get_pattern_fill().get_foreground_color());
        TS_ASSERT(const_fill.get_pattern_fill().get_background_color());
    }
    
    void test_hash()
    {
        xlnt::fill pattern_fill = xlnt::fill::pattern(xlnt::pattern_fill::type::solid);
        xlnt::fill gradient_fill_linear = xlnt::fill::gradient(xlnt::gradient_fill::type::linear);
        xlnt::fill gradient_fill_path = xlnt::fill::gradient(xlnt::gradient_fill::type::path);

        TS_ASSERT_DIFFERS(pattern_fill.hash(), gradient_fill_linear.hash());
        TS_ASSERT_DIFFERS(gradient_fill_linear.hash(), gradient_fill_path.hash());
        TS_ASSERT_DIFFERS(gradient_fill_path.hash(), pattern_fill.hash());
    }
};
