#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/xlnt.hpp>
#include "helpers/path_helper.hpp"

class test_theme : public CxxTest::TestSuite
{
public:
    void test_write_theme()
    {
        auto content = xlnt::writer::write_theme();
        
        std::string comparison_file = PathHelper::GetDataDirectory() + "/writer/expected/theme1.xml";
        std::ifstream t(comparison_file);
        std::stringstream buffer;
        buffer << t.rdbuf();
        std::string expected = buffer.str();
        
        TS_ASSERT_EQUALS(buffer.str(), content);
    }
};
