// Copyright (c) 2014-2018 Thomas Fussell
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

#include <iostream>

#include <helpers/test_suite.hpp>

#include <xlnt/styles/color.hpp>

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
        const std::vector<std::pair<xlnt::color, std::string>> known_colors{
            {xlnt::color::black(), "FF000000"},
            {xlnt::color::white(), "FFFFFFFF"},
            {xlnt::color::red(), "FFFF0000"},
            {xlnt::color::darkred(), "FF8B0000"},
            {xlnt::color::blue(), "FF0000FF"},
            {xlnt::color::darkblue(), "FF00008B"},
            {xlnt::color::green(), "FF00FF00"},
            {xlnt::color::darkgreen(), "FF008B00"},
            {xlnt::color::yellow(), "FFFFFF00"},
            {xlnt::color::darkyellow(), "FFCCCC00"}};

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
        indexed.indexed().index(2);
        xlnt_assert_equals(indexed.indexed().index(), 2);
        xlnt_assert_throws(indexed.theme(), xlnt::invalid_attribute);
        xlnt_assert_throws(indexed.rgb(), xlnt::invalid_attribute);

        xlnt::color theme = xlnt::theme_color(3);
        xlnt_assert(!theme.auto_());
        xlnt_assert_equals(theme.theme().index(), 3);
        theme.theme().index(4);
        xlnt_assert_equals(theme.theme().index(), 4);
        xlnt_assert_throws(theme.indexed(), xlnt::invalid_attribute);
        xlnt_assert_throws(theme.rgb(), xlnt::invalid_attribute);
    }
};
static color_test_suite x;