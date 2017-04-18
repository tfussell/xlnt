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

#include <ctime>
#include <iostream>
#include <sstream>

#include <helpers/test_suite.hpp>
#include <xlnt/xlnt.hpp>

class rich_text_test_suite : public test_suite
{
public:
    rich_text_test_suite()
    {
        register_test(test_operators);
    }

    void test_operators()
    {
        xlnt::rich_text text1;
        xlnt::rich_text text2;
        xlnt_assert_equals(text1, text2);
        xlnt::rich_text_run run_default;
        text1.add_run(run_default);
        xlnt_assert_differs(text1, text2);
        text2.add_run(run_default);
        xlnt_assert_equals(text1, text2);

        xlnt::rich_text_run run_formatted;
        xlnt::font run_font;
        run_font.color(xlnt::color::green());
        run_font.name("Cambria");
        run_font.scheme("ascheme");
        run_font.size(40);
        run_font.family(17);
        run_formatted.second = run_font;

        xlnt::rich_text text_formatted;
        text_formatted.add_run(run_formatted);

        xlnt::rich_text_run run_color_differs = run_formatted;
        run_font = xlnt::font();
        run_font.color(xlnt::color::red());
        run_color_differs.second = run_font;
        xlnt::rich_text text_color_differs;
        text_color_differs.add_run(run_color_differs);
        xlnt_assert_differs(text_formatted, text_color_differs);

        xlnt::rich_text_run run_font_differs = run_formatted;
        run_font = xlnt::font();
        run_font.name("Calibri");
        run_font_differs.second = run_font;
        xlnt::rich_text text_font_differs;
        text_font_differs.add_run(run_font_differs);
        xlnt_assert_differs(text_formatted, text_font_differs);

        xlnt::rich_text_run run_scheme_differs = run_formatted;
        run_font = xlnt::font();
        run_font.scheme("bscheme");
        run_scheme_differs.second = run_font;
        xlnt::rich_text text_scheme_differs;
        text_scheme_differs.add_run(run_scheme_differs);
        xlnt_assert_differs(text_formatted, text_scheme_differs);

        xlnt::rich_text_run run_size_differs = run_formatted;
        run_font = xlnt::font();
        run_font.size(41);
        run_size_differs.second = run_font;
        xlnt::rich_text text_size_differs;
        text_size_differs.add_run(run_size_differs);
        xlnt_assert_differs(text_formatted, text_size_differs);

        xlnt::rich_text_run run_family_differs = run_formatted;
        run_font = xlnt::font();
        run_font.family(18);
        run_family_differs.second = run_font;
        xlnt::rich_text text_family_differs;
        text_family_differs.add_run(run_family_differs);
        xlnt_assert_differs(text_formatted, text_family_differs);
    }
};
