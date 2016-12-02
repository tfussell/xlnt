#pragma once

#include <ctime>
#include <iostream>
#include <sstream>
#include <cxxtest/TestSuite.h>

#include <xlnt/xlnt.hpp>

class test_formatted_text : public CxxTest::TestSuite
{
public:
    void test_operators()
    {
        xlnt::formatted_text text1;
        xlnt::formatted_text text2;
        TS_ASSERT_EQUALS(text1, text2);
        xlnt::text_run run_default;
        text1.add_run(run_default);
        TS_ASSERT_DIFFERS(text1, text2);
        text2.add_run(run_default);
        TS_ASSERT_EQUALS(text1, text2);

        xlnt::text_run run_formatted;
        run_formatted.color(xlnt::color::green());
        run_formatted.font("Cambria");
        run_formatted.scheme("ascheme");
        run_formatted.size(40);
        run_formatted.family(17);

        xlnt::formatted_text text_formatted;
        text_formatted.add_run(run_formatted);

        xlnt::text_run run_color_differs = run_formatted;
        run_color_differs.color(xlnt::color::red());
        xlnt::formatted_text text_color_differs;
        text_color_differs.add_run(run_color_differs);
        TS_ASSERT_DIFFERS(text_formatted, text_color_differs);

        xlnt::text_run run_font_differs = run_formatted;
        run_font_differs.font("Calibri");
        xlnt::formatted_text text_font_differs;
        text_font_differs.add_run(run_font_differs);
        TS_ASSERT_DIFFERS(text_formatted, text_font_differs);

        xlnt::text_run run_scheme_differs = run_formatted;
        run_scheme_differs.scheme("bscheme");
        xlnt::formatted_text text_scheme_differs;
        text_scheme_differs.add_run(run_scheme_differs);
        TS_ASSERT_DIFFERS(text_formatted, text_scheme_differs);

        xlnt::text_run run_size_differs = run_formatted;
        run_size_differs.size(41);
        xlnt::formatted_text text_size_differs;
        text_size_differs.add_run(run_size_differs);
        TS_ASSERT_DIFFERS(text_formatted, text_size_differs);

        xlnt::text_run run_family_differs = run_formatted;
        run_family_differs.family(18);
        xlnt::formatted_text text_family_differs;
        text_family_differs.add_run(run_family_differs);
        TS_ASSERT_DIFFERS(text_formatted, text_family_differs);
    }
};
