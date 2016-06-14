#pragma once

#include <cxxtest/TestSuite.h>

#include <xlnt/utils/utf8string.hpp>

class test_utf8string : public CxxTest::TestSuite
{
public:
    void test_utf8()
    {
        auto utf8_valid = xlnt::utf8string::from_utf8("abc");
        auto utf8_invalid = xlnt::utf8string::from_utf8("\xc3\x28");
        
        TS_ASSERT(utf8_valid.is_valid());
        TS_ASSERT(!utf8_invalid.is_valid());
    }

    void test_latin1()
    {
        auto latin1_valid = xlnt::utf8string::from_latin1("abc");
        TS_ASSERT(latin1_valid.is_valid());
    }
    
    void test_utf16()
    {
        auto utf16_valid = xlnt::utf8string::from_utf16("abc");
        TS_ASSERT(utf16_valid.is_valid());
    }
    
    void test_utf32()
    {
        auto utf32_valid = xlnt::utf8string::from_utf32("abc");
        TS_ASSERT(utf32_valid.is_valid());
    }
};
