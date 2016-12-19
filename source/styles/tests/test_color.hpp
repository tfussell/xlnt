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
            TS_ASSERT_EQUALS(pair.first.rgb().hex_string(), pair.second);
        }
    }
    
    void test_non_rgb_colors()
    {
		xlnt::color indexed = xlnt::indexed_color(1);
		TS_ASSERT(!indexed.is_auto());
        TS_ASSERT_EQUALS(indexed.indexed().index(), 1);
        TS_ASSERT_THROWS(indexed.theme(), xlnt::invalid_attribute);
        TS_ASSERT_THROWS(indexed.rgb(), xlnt::invalid_attribute);

		xlnt::color theme = xlnt::theme_color(3);
		TS_ASSERT(!theme.is_auto());
        TS_ASSERT_EQUALS(theme.theme().index(), 3);
        TS_ASSERT_THROWS(theme.indexed(), xlnt::invalid_attribute);
        TS_ASSERT_THROWS(theme.rgb(), xlnt::invalid_attribute);
    }
};
