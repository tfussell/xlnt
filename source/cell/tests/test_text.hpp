#pragma once

#include <ctime>
#include <iostream>
#include <sstream>
#include <cxxtest/TestSuite.h>

#include <xlnt/xlnt.hpp>

class test_rich_text : public CxxTest::TestSuite
{
public:
    void test_operators()
    {
        xlnt::rich_text text1;
        xlnt::rich_text text2;
        TS_ASSERT_EQUALS(text1, text2);
        xlnt::rich_text_run run_default;
        text1.add_run(run_default);
        TS_ASSERT_DIFFERS(text1, text2);
        text2.add_run(run_default);
        TS_ASSERT_EQUALS(text1, text2);

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
        TS_ASSERT_DIFFERS(text_formatted, text_color_differs);

        xlnt::rich_text_run run_font_differs = run_formatted;
        run_font = xlnt::font();
        run_font.name("Calibri");
        run_font_differs.second = run_font;
        xlnt::rich_text text_font_differs;
        text_font_differs.add_run(run_font_differs);
        TS_ASSERT_DIFFERS(text_formatted, text_font_differs);

        xlnt::rich_text_run run_scheme_differs = run_formatted;
        run_font = xlnt::font();
        run_font.scheme("bscheme");
        run_scheme_differs.second = run_font;
        xlnt::rich_text text_scheme_differs;
        text_scheme_differs.add_run(run_scheme_differs);
        TS_ASSERT_DIFFERS(text_formatted, text_scheme_differs);

        xlnt::rich_text_run run_size_differs = run_formatted;
        run_font = xlnt::font();
        run_font.size(41);
        run_size_differs.second = run_font;
        xlnt::rich_text text_size_differs;
        text_size_differs.add_run(run_size_differs);
        TS_ASSERT_DIFFERS(text_formatted, text_size_differs);

        xlnt::rich_text_run run_family_differs = run_formatted;
        run_font = xlnt::font();
        run_font.family(18);
        run_family_differs.second = run_font;
        xlnt::rich_text text_family_differs;
        text_family_differs.add_run(run_family_differs);
        TS_ASSERT_DIFFERS(text_formatted, text_family_differs);
    }
};
