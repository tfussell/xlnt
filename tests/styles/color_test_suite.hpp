// Copyright (c) 2014-2017 Thomas Fussell
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, WRISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE
//
// @license: http://www.opensource.org/licenses/mit-license.php
// @author: see AUTHORS file

#pragma once

#include <iostream>

#include <helpers/test_suite.hpp>
#include <xlnt/xlnt.hpp>

class color_test_suite : public test_suite
{
public:
    color_test_suite()
    {
        register_test(test_known_colors);
        register_test(test_non_rgb_colors);
    }

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
            xlnt_assert_equals(pair.first.rgb().hex_string(), pair.second);
        }
    }
    
    void test_non_rgb_colors()
    {
		xlnt::color indexed = xlnt::indexed_color(1);
		xlnt_assert(!indexed.auto_());
        xlnt_assert_equals(indexed.indexed().index(), 1);
        xlnt_assert_throws(indexed.theme(), xlnt::invalid_attribute);
        xlnt_assert_throws(indexed.rgb(), xlnt::invalid_attribute);

		xlnt::color theme = xlnt::theme_color(3);
		xlnt_assert(!theme.auto_());
        xlnt_assert_equals(theme.theme().index(), 3);
        xlnt_assert_throws(theme.indexed(), xlnt::invalid_attribute);
        xlnt_assert_throws(theme.rgb(), xlnt::invalid_attribute);
    }
};
