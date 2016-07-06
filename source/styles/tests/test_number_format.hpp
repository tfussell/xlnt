#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include "pugixml.hpp"
#include <xlnt/xlnt.hpp>

class test_number_format : public CxxTest::TestSuite
{
public:
    void test_simple_format()
    {
        xlnt::number_format nf;

        nf.set_format_string("\"positive\"General;\"negative\"General");
        auto formatted = nf.format(3.14, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "positive3.14");
        formatted = nf.format(-3.14, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "negative3.14");

        nf.set_format_string("\"any\"General");
        formatted = nf.format(-3.14, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "-any3.14");
    }

    void test_simple_date()
    {
        auto date = xlnt::date(2016, 6, 18);
        auto date_number = date.to_number(xlnt::calendar::windows_1900);

        xlnt::number_format nf = xlnt::number_format::date_ddmmyyyy();
        auto formatted = nf.format(date_number, xlnt::calendar::windows_1900);

        TS_ASSERT_EQUALS(formatted, "18/06/16");
    }

    void test_short_month()
    {
        auto date = xlnt::date(2016, 6, 18);
        auto date_number = date.to_number(xlnt::calendar::windows_1900);

        xlnt::number_format nf;
        nf.set_format_string("m");
        auto formatted = nf.format(date_number, xlnt::calendar::windows_1900);

        TS_ASSERT_EQUALS(formatted, "6");
    }

    void test_month_abbreviation()
    {
        auto date = xlnt::date(2016, 6, 18);
        auto date_number = date.to_number(xlnt::calendar::windows_1900);

        xlnt::number_format nf;
        nf.set_format_string("mmm");
        auto formatted = nf.format(date_number, xlnt::calendar::windows_1900);

        TS_ASSERT_EQUALS(formatted, "Jun");
    }

    void test_month_name()
    {
        auto date = xlnt::date(2016, 6, 18);
        auto date_number = date.to_number(xlnt::calendar::windows_1900);

        xlnt::number_format nf;
        nf.set_format_string("mmmm");
        auto formatted = nf.format(date_number, xlnt::calendar::windows_1900);

        TS_ASSERT_EQUALS(formatted, "June");
    }

    void test_short_day()
    {
        auto date = xlnt::date(2016, 6, 8);
        auto date_number = date.to_number(xlnt::calendar::windows_1900);

        xlnt::number_format nf;
        nf.set_format_string("d");
        auto formatted = nf.format(date_number, xlnt::calendar::windows_1900);

        TS_ASSERT_EQUALS(formatted, "8");
    }

    void test_long_day()
    {
        auto date = xlnt::date(2016, 6, 8);
        auto date_number = date.to_number(xlnt::calendar::windows_1900);

        xlnt::number_format nf;
        nf.set_format_string("dd");
        auto formatted = nf.format(date_number, xlnt::calendar::windows_1900);

        TS_ASSERT_EQUALS(formatted, "08");
    }

    void test_long_year()
    {
        auto date = xlnt::date(2016, 6, 18);
        auto date_number = date.to_number(xlnt::calendar::windows_1900);

        xlnt::number_format nf;
        nf.set_format_string("yyyy");
        auto formatted = nf.format(date_number, xlnt::calendar::windows_1900);

        TS_ASSERT_EQUALS(formatted, "2016");
    }

    void test_time_24_hour()
    {
        auto time = xlnt::time(20, 15, 10);
        auto time_number = time.to_number();

        xlnt::number_format nf = xlnt::number_format::date_time4();
        auto formatted = nf.format(time_number, xlnt::calendar::windows_1900);

        TS_ASSERT_EQUALS(formatted, "20:15:10");
    }

    void test_time_12_hour_am()
    {
        auto time = xlnt::time(8, 15, 10);
        auto time_number = time.to_number();

        xlnt::number_format nf = xlnt::number_format::date_time2();
        auto formatted = nf.format(time_number, xlnt::calendar::windows_1900);

        TS_ASSERT_EQUALS(formatted, "8:15:10 AM");
    }

    void test_time_12_hour_pm()
    {
        auto time = xlnt::time(20, 15, 10);
        auto time_number = time.to_number();

        xlnt::number_format nf = xlnt::number_format::date_time2();
        auto formatted = nf.format(time_number, xlnt::calendar::windows_1900);

        TS_ASSERT_EQUALS(formatted, "8:15:10 PM");
    }

    void test_long_hour_12_hour()
    {
        auto time = xlnt::time(20, 15, 10);
        auto time_number = time.to_number();

        xlnt::number_format nf;
        nf.set_format_string("hh AM/PM");
        auto formatted = nf.format(time_number, xlnt::calendar::windows_1900);

        TS_ASSERT_EQUALS(formatted, "08 PM");
    }

    void test_long_hour_24_hour()
    {
        auto time = xlnt::time(20, 15, 10);
        auto time_number = time.to_number();

        xlnt::number_format nf;
        nf.set_format_string("hh");
        auto formatted = nf.format(time_number, xlnt::calendar::windows_1900);

        TS_ASSERT_EQUALS(formatted, "20");
    }

    void test_short_minute()
    {
        auto time = xlnt::time(20, 5, 10);
        auto time_number = time.to_number();

        xlnt::number_format nf;
        nf.set_format_string("h:m");
        auto formatted = nf.format(time_number, xlnt::calendar::windows_1900);

        TS_ASSERT_EQUALS(formatted, "20:5");
    }

    void test_long_minute()
    {
        auto time = xlnt::time(20, 5, 10);
        auto time_number = time.to_number();

        xlnt::number_format nf;
        nf.set_format_string("h:mm");
        auto formatted = nf.format(time_number, xlnt::calendar::windows_1900);

        TS_ASSERT_EQUALS(formatted, "20:05");
    }

    void test_short_second()
    {
        auto time = xlnt::time(20, 15, 1);
        auto time_number = time.to_number();

        xlnt::number_format nf;
        nf.set_format_string("h:m:s");
        auto formatted = nf.format(time_number, xlnt::calendar::windows_1900);

        TS_ASSERT_EQUALS(formatted, "20:15:1");
    }

    void test_long_second()
    {
        auto time = xlnt::time(20, 15, 1);
        auto time_number = time.to_number();

        xlnt::number_format nf;
        nf.set_format_string("h:m:ss");
        auto formatted = nf.format(time_number, xlnt::calendar::windows_1900);

        TS_ASSERT_EQUALS(formatted, "20:15:01");
    }

    void test_trailing_space()
    {
        auto time = xlnt::time(20, 15, 1);
        auto time_number = time.to_number();

        xlnt::number_format nf;
        nf.set_format_string("h:m:ss ");
        auto formatted = nf.format(time_number, xlnt::calendar::windows_1900);

        TS_ASSERT_EQUALS(formatted, "20:15:01 ");
    }

    void test_text_section_string()
    {
        xlnt::number_format nf;
        nf.set_format_string("General;General;General;[Green]\"a\"@\"b\"");

        auto formatted = nf.format("text");

        TS_ASSERT_EQUALS(formatted, "atextb");
    }

    void test_text_section_no_string()
    {
        xlnt::number_format nf;
        nf.set_format_string("General;General;General;[Green]\"ab\"");

        auto formatted = nf.format("text");

        TS_ASSERT_EQUALS(formatted, "ab");
    }

    void test_conditional_format()
    {
        xlnt::number_format nf;

        nf.set_format_string("[>5]General\"first\";[>3]\"second\"General;\"third\"General");
        auto formatted = nf.format(6, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "6first");
        formatted = nf.format(4, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "second4");
        formatted = nf.format(5, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "second5");
        formatted = nf.format(3, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "third3");
        formatted = nf.format(2, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "third2");

        nf.set_format_string("[>=5]\"first\"General;[>=3]\"second\"General;\"third\"General");
        formatted = nf.format(5, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "first5");
        formatted = nf.format(6, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "first6");
        formatted = nf.format(4, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "second4");
        formatted = nf.format(3, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "second3");
        formatted = nf.format(2, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "third2");

        nf.set_format_string("[<1]\"first\"General;[<5]\"second\"General;\"third\"General");
        formatted = nf.format(0, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "first0");
        formatted = nf.format(1, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "second1");
        formatted = nf.format(5, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "third5");
        formatted = nf.format(6, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "third6");

        nf.set_format_string("[<=1]\"first\"General;[<=5]\"second\"General;\"third\"General");
        formatted = nf.format(-1000, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "-first1000");
        formatted = nf.format(0, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "first0");
        formatted = nf.format(1, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "first1");
        formatted = nf.format(4, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "second4");
        formatted = nf.format(5, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "second5");
        formatted = nf.format(6, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "third6");
        formatted = nf.format(1000, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "third1000");

        nf.set_format_string("[=1]\"first\"General;[=2]\"second\"General;\"third\"General");
        formatted = nf.format(1, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "first1");
        formatted = nf.format(2, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "second2");
        formatted = nf.format(3, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "third3");
        formatted = nf.format(0, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "third0");
    }
    
    void test_locale_currency()
    {
        xlnt::number_format nf;

        nf.set_format_string("[$€-407]#,##0.00");
        auto formatted = nf.format(-45000.1, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "-€45,000.10");

        nf.set_format_string("[$$-1009]#,##0.00");
        formatted = nf.format(-45000.1, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "-$45,000.10");
    }
    
    void test_bad_country()
    {
        xlnt::number_format nf;

        nf.set_format_string("[$-]#,##0.00");
        TS_ASSERT_THROWS(nf.format(1.2, xlnt::calendar::windows_1900), std::runtime_error);

        nf.set_format_string("[$-G]#,##0.00");
        TS_ASSERT_THROWS(nf.format(1.2, xlnt::calendar::windows_1900), std::runtime_error);

        nf.set_format_string("[$-4002]#,##0.00");
        TS_ASSERT_THROWS(nf.format(1.2, xlnt::calendar::windows_1900), std::runtime_error);

        nf.set_format_string("[-4001]#,##0.00");
        TS_ASSERT_THROWS(nf.format(1.2, xlnt::calendar::windows_1900), std::runtime_error);
    }

    void test_duplicate_bracket_sections()
    {
        xlnt::number_format nf;

        nf.set_format_string("[Red][Green]#,##0.00");
        TS_ASSERT_THROWS(nf.format(1.2, xlnt::calendar::windows_1900), std::runtime_error);

        nf.set_format_string("[$-403][$-4001]#,##0.00");
        TS_ASSERT_THROWS(nf.format(1.2, xlnt::calendar::windows_1900), std::runtime_error);

        nf.set_format_string("[>3][>4]#,##0.00");
        TS_ASSERT_THROWS(nf.format(1.2, xlnt::calendar::windows_1900), std::runtime_error);
    }
    
    void test_bad_format()
    {
        xlnt::number_format nf;

        nf.set_format_string("[=1]\"first\"General;[=2]\"second\"General;[=3]\"third\"General");
        TS_ASSERT_THROWS(nf.format(1.2, xlnt::calendar::windows_1900), std::runtime_error);

        nf.set_format_string("\"first\"General;\"second\"General;\"third\"General;\"fourth\"General;\"fifth\"General");
        TS_ASSERT_THROWS(nf.format(1.2, xlnt::calendar::windows_1900), std::runtime_error);
    }
    
    void format_and_test(const xlnt::number_format &nf, const std::array<std::string, 4> &expect)
    {
        long double positive = 42503.1234;
        long double negative = -1 * positive;
        long double zero = 0;
        const std::string text = "text";
        xlnt::calendar calendar = xlnt::calendar::windows_1900;

        TS_ASSERT_EQUALS(nf.format(positive, calendar), expect[0]);
        TS_ASSERT_EQUALS(nf.format(negative, calendar), expect[1]);
        TS_ASSERT_EQUALS(nf.format(zero, calendar), expect[2]);
        TS_ASSERT_EQUALS(nf.format(text), expect[3]);
    }

    // General
    void test_builtin_format_0()
    {
        format_and_test(xlnt::number_format::general(), {{"42503.1234", "-42503.1234", "0", "text"}});
    }

    // 0
    void test_builtin_format_1()
    {
        format_and_test(xlnt::number_format::number(), {{"42503", "-42503", "0", "text"}});
    }
    // 0.00
    void test_builtin_format_2()
    {
        format_and_test(xlnt::number_format::number_00(), {{"42503.12", "-42503.12", "0.00", "text"}});
    }

    // #,##0
    void test_builtin_format_3()
    {
        format_and_test(xlnt::number_format::from_builtin_id(3), {{"42,503", "-42,503", "0", "text"}});
    }

    // #,##0.00
    void test_builtin_format_4()
    {
        format_and_test(xlnt::number_format::from_builtin_id(4), {{"42,503.12", "-42,503.12", "0.00", "text"}});
    }

    // #,##0;-#,##0
    void test_builtin_format_5()
    {
        format_and_test(xlnt::number_format::from_builtin_id(5), {{"42,503", "-42,503", "0", "text"}});
    }

    // #,##0;[Red]-#,##0
    void test_builtin_format_6()
    {
        format_and_test(xlnt::number_format::from_builtin_id(6), {{"42,503", "-42,503", "0", "text"}});
    }

    // #,##0.00;-#,##0.00
    void test_builtin_format_7()
    {
        format_and_test(xlnt::number_format::from_builtin_id(7), {{"42,503.12", "-42,503.12", "0.00", "text"}});
    }

    // #,##0.00;[Red]-#,##0.00
    void test_builtin_format_8()
    {
        format_and_test(xlnt::number_format::from_builtin_id(8), {{"42,503.12", "-42,503.12", "0.00", "text"}});
    }

    // 0%
    void test_builtin_format_9()
    {
        format_and_test(xlnt::number_format::percentage(), {{"4250312%", "-4250312%", "0%", "text"}});
    }

    // 0.00%
    void test_builtin_format_10()
    {
        format_and_test(xlnt::number_format::percentage_00(), {{"4250312.34%", "-4250312.34%", "0.00%", "text"}});
    }

    // 0.00E+00
    void test_builtin_format_11()
    {
        format_and_test(xlnt::number_format::from_builtin_id(11), {{"4.25E+04", "-4.25E+04", "0.00E+00", "text"}});
    }

    // # ?/?
    void _test_builtin_format_12()
    {
        format_and_test(xlnt::number_format::from_builtin_id(12), {{"42503 1/8", "-42503 1/8", "0", "text"}});
    }

    // # ??/??
    void _test_builtin_format_13()
    {
        format_and_test(xlnt::number_format::from_builtin_id(13), {{"42503 10/81", "-42503 10/81", "0", "text"}});
    }

    // mm-dd-yy
    void _test_builtin_format_14()
    {
        format_and_test(xlnt::number_format::from_builtin_id(14), {{"05-13-16", "###########", "01-00-00", "text"}});
    }

    // d-mmm-yy
    void _test_builtin_format_15()
    {
        format_and_test(xlnt::number_format::from_builtin_id(15), {{"13-May-16", "###########", "0-Jan-00", "text"}});
    }

    // d-mmm
    void _test_builtin_format_16()
    {
        format_and_test(xlnt::number_format::from_builtin_id(16), {{"13-May", "###########", "0-Jan", "text"}});
    }

    // mmm-yy
    void _test_builtin_format_17()
    {
        format_and_test(xlnt::number_format::from_builtin_id(17), {{"May-16", "###########", "Jan-00", "text"}});
    }

    // h:mm AM/PM
    void _test_builtin_format_18()
    {
        format_and_test(xlnt::number_format::from_builtin_id(18), {{"2:57 AM", "###########", "12:00 AM", "text"}});
    }
    
    // h:mm:ss AM/PM
    void _test_builtin_format_19()
    {
        format_and_test(xlnt::number_format::from_builtin_id(19), {{"2:57:42 AM", "###########", "12:00:00 AM", "text"}});
    }
    
    // h:mm
    void _test_builtin_format_20()
    {
        format_and_test(xlnt::number_format::from_builtin_id(20), {{"2:57", "###########", "0:00", "text"}});
    }
    
    // h:mm:ss
    void _test_builtin_format_21()
    {
        format_and_test(xlnt::number_format::from_builtin_id(21), {{"2:57:42", "###########", "0:00:00", "text"}});
    }
    
    // m/d/yy h:mm
    void _test_builtin_format_22()
    {
        format_and_test(xlnt::number_format::from_builtin_id(22), {{"5/13/16 2:57", "###########", "1/0/00 0:00", "text"}});
    }
};
