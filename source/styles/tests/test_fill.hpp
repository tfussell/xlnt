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
        
        TS_ASSERT_EQUALS(fill.get_gradient_type(), xlnt::fill::gradient_type::none);
        fill.set_gradient_type(xlnt::fill::gradient_type::linear);
        TS_ASSERT_EQUALS(fill.get_gradient_type(), xlnt::fill::gradient_type::linear);
        
        TS_ASSERT(!fill.get_foreground_color());
        fill.set_foreground_color(xlnt::color::black());
        TS_ASSERT(fill.get_foreground_color());
        
        TS_ASSERT(!fill.get_background_color());
        fill.set_background_color(xlnt::color::black());
        TS_ASSERT(fill.get_background_color());
        
        TS_ASSERT(!fill.get_start_color());
        fill.set_start_color(xlnt::color::black());
        TS_ASSERT(fill.get_start_color());
        
        TS_ASSERT(!fill.get_end_color());
        fill.set_end_color(xlnt::color::black());
        TS_ASSERT(fill.get_end_color());
        
        TS_ASSERT_EQUALS(fill.get_rotation(), 0);
        fill.set_rotation(1);
        TS_ASSERT_EQUALS(fill.get_rotation(), 1);
        
        TS_ASSERT_EQUALS(fill.get_gradient_bottom(), 0);
        TS_ASSERT_EQUALS(fill.get_gradient_left(), 0);
        TS_ASSERT_EQUALS(fill.get_gradient_top(), 0);
        TS_ASSERT_EQUALS(fill.get_gradient_right(), 0);
        
        const auto &const_fill = fill;
        TS_ASSERT(const_fill.get_foreground_color());
        TS_ASSERT(const_fill.get_background_color());
        TS_ASSERT(const_fill.get_start_color());
        TS_ASSERT(const_fill.get_end_color());
    }
    
    void test_hash()
    {
        xlnt::fill none_fill;

        xlnt::fill solid_fill;
        solid_fill.set_type(xlnt::fill::type::solid);

        xlnt::fill pattern_fill;
        pattern_fill.set_type(xlnt::fill::type::pattern);

        xlnt::fill gradient_fill_linear;
        gradient_fill_linear.set_type(xlnt::fill::type::gradient);
        gradient_fill_linear.set_gradient_type(xlnt::fill::gradient_type::linear);
        
        xlnt::fill gradient_fill_path;
        gradient_fill_path.set_type(xlnt::fill::type::gradient);
        gradient_fill_path.set_gradient_type(xlnt::fill::gradient_type::linear);
        gradient_fill_path.set_start_color(xlnt::color::red());
        gradient_fill_path.set_end_color(xlnt::color::green());

        TS_ASSERT_DIFFERS(none_fill.hash(), solid_fill.hash());
        TS_ASSERT_DIFFERS(solid_fill.hash(), pattern_fill.hash());
        TS_ASSERT_DIFFERS(pattern_fill.hash(), gradient_fill_linear.hash());
        TS_ASSERT_DIFFERS(gradient_fill_linear.hash(), gradient_fill_path.hash());
        TS_ASSERT_DIFFERS(gradient_fill_path.hash(), none_fill.hash());
    }
};
