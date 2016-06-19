#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/xlnt.hpp>

class test_color : public CxxTest::TestSuite
{
public:
    void test_known_colors()
    {
        const std::vector<std::pair<xlnt::color, std::string>> known_colors
        {
            { xlnt::color::black(), "ff000000" },
            { xlnt::color::white(), "ffffffff" },
            { xlnt::color::red(), "ffff0000" },
            { xlnt::color::darkred(), "ff8b0000" },
            { xlnt::color::blue(), "ff0000ff" },
            { xlnt::color::darkblue(), "ff00008b" },
            { xlnt::color::green(), "ff00ff00" },
            { xlnt::color::darkgreen(), "ff008b00" },
            { xlnt::color::yellow(), "ffffff00" },
            { xlnt::color::darkyellow(), "ffcccc00" }
        };
        
        for (auto pair : known_colors)
        {
            TS_ASSERT_EQUALS(pair.first.get_rgb_string(), pair.second);
        }
    }
    
    void test_non_rgb_colors()
    {
        xlnt::color indexed;
        indexed.set_index(1);
        TS_ASSERT_EQUALS(indexed.get_index(), 1);
        TS_ASSERT_THROWS(indexed.get_auto(), std::runtime_error);
        TS_ASSERT_THROWS(indexed.get_theme(), std::runtime_error);
        TS_ASSERT_THROWS(indexed.get_rgb_string(), std::runtime_error);

        xlnt::color auto_;
        auto_.set_auto(2);
        TS_ASSERT_EQUALS(auto_.get_auto(), 2);
        TS_ASSERT_THROWS(auto_.get_theme(), std::runtime_error);
        TS_ASSERT_THROWS(auto_.get_index(), std::runtime_error);
        TS_ASSERT_THROWS(auto_.get_rgb_string(), std::runtime_error);

        xlnt::color theme;
        theme.set_theme(3);
        TS_ASSERT_EQUALS(theme.get_theme(), 3);
        TS_ASSERT_THROWS(theme.get_auto(), std::runtime_error);
        TS_ASSERT_THROWS(theme.get_index(), std::runtime_error);
        TS_ASSERT_THROWS(theme.get_rgb_string(), std::runtime_error);
    }
};
