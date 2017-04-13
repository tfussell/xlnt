#pragma once

#include <iostream>
#include <helpers/test_suite.hpp>

#include <xlnt/xlnt.hpp>

class test_color : public test_suite
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
            assert_equals(pair.first.rgb().hex_string(), pair.second);
        }
    }
    
    void test_non_rgb_colors()
    {
		xlnt::color indexed = xlnt::indexed_color(1);
		assert(!indexed.auto_());
        assert_equals(indexed.indexed().index(), 1);
        assert_throws(indexed.theme(), xlnt::invalid_attribute);
        assert_throws(indexed.rgb(), xlnt::invalid_attribute);

		xlnt::color theme = xlnt::theme_color(3);
		assert(!theme.auto_());
        assert_equals(theme.theme().index(), 3);
        assert_throws(theme.indexed(), xlnt::invalid_attribute);
        assert_throws(theme.rgb(), xlnt::invalid_attribute);
    }
};
