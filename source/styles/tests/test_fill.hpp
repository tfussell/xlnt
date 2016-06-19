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
    }
    
    void test_hash()
    {
        xlnt::fill none_fill;

        xlnt::fill solid_fill;
        solid_fill.set_type(xlnt::fill::type::solid);

        xlnt::fill pattern_fill;
        pattern_fill.set_type(xlnt::fill::type::pattern);

        xlnt::fill gradient_fill;
        gradient_fill.set_type(xlnt::fill::type::gradient);

        TS_ASSERT_DIFFERS(none_fill.hash(), solid_fill.hash());
        TS_ASSERT_DIFFERS(solid_fill.hash(), pattern_fill.hash());
        TS_ASSERT_DIFFERS(pattern_fill.hash(), gradient_fill.hash());
        TS_ASSERT_DIFFERS(gradient_fill.hash(), none_fill.hash());
    }
};
