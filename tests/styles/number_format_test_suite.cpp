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

#include <xlnt/styles/number_format.hpp>
#include <xlnt/utils/date.hpp>
#include <xlnt/utils/time.hpp>
#include <xlnt/utils/timedelta.hpp>

class number_format_test_suite : public test_suite
{
public:
    number_format_test_suite()
    {
        register_test(test_basic);
        register_test(test_simple_format);
        register_test(test_simple_date);
        register_test(test_short_month);
        register_test(test_month_abbreviation);
        register_test(test_month_name);
        register_test(test_basic);
        register_test(test_simple_format);
        register_test(test_upper_case_date);
        register_test(test_simple_date);
        register_test(test_short_day);
        register_test(test_long_day);
        register_test(test_long_year);
        register_test(test_day_name);
        register_test(test_day_abbreviation);
        register_test(test_month_letter);
        register_test(test_time_24_hour);
        register_test(test_elapsed_minutes);
        register_test(test_second_fractional_leading_zero);
        register_test(test_second_fractional);
        register_test(test_elapsed_seconds);
        register_test(test_time_12_hour_am);
        register_test(test_time_12_hour_pm);
        register_test(test_long_hour_12_hour);
        register_test(test_long_hour_12_hour_ap);
        register_test(test_long_hour_24_hour);
        register_test(test_short_minute);
        register_test(test_long_minute);
        register_test(test_short_second);
        register_test(test_long_second);
        register_test(test_trailing_space);
        register_test(test_text_section_string);
        register_test(test_text_section_no_string);
        register_test(test_text_section_no_text);
        register_test(test_conditional_format);
        register_test(test_space);
        register_test(test_fill);
        register_test(test_placeholders_zero);
        register_test(test_placeholders_space);
        register_test(test_scientific);
        register_test(test_locale_currency);
        register_test(test_bad_country);
        register_test(test_duplicate_bracket_sections);
        register_test(test_escaped_quote_string);
        register_test(test_thousands_scale);
        register_test(test_colors);
        register_test(test_bad_format);
        register_test(test_builtin_format_0);
        register_test(test_builtin_format_1);
        register_test(test_builtin_format_2);
        register_test(test_builtin_format_3);
        register_test(test_builtin_format_4);
        register_test(test_builtin_format_9);
        register_test(test_builtin_format_10);
        register_test(test_builtin_format_11);
        register_test(test_builtin_format_12);
        register_test(test_builtin_format_13);
        register_test(test_builtin_format_14);
        register_test(test_builtin_format_15);
        register_test(test_builtin_format_16);
        register_test(test_builtin_format_17);
        register_test(test_builtin_format_18);
        register_test(test_builtin_format_19);
        register_test(test_builtin_format_20);
        register_test(test_builtin_format_21);
        register_test(test_builtin_format_22);
        register_test(test_builtin_format_37);
        register_test(test_builtin_format_38);
        register_test(test_builtin_format_39);
        register_test(test_builtin_format_40);
        register_test(test_builtin_format_45);
        register_test(test_builtin_format_46);
        register_test(test_builtin_format_47);
        register_test(test_builtin_format_48);
        register_test(test_builtin_format_49);
        register_test(test_builtin_format_date_yyyymmdd);
        register_test(test_builtin_format_date_dmyslash);
        register_test(test_builtin_format_date_dmyminus);
        register_test(test_builtin_format_date_dmminus);
        register_test(test_builtin_format_date_myminus);
    }

    void test_basic()
    {
        xlnt::number_format no_id("#\\x\\y\\z");
        xlnt_assert_throws(no_id.id(), std::runtime_error);

        xlnt::number_format id("General", 200);
        xlnt_assert_equals(id.id(), 200);
        xlnt_assert_equals(id.format_string(), "General");

        xlnt::number_format general(0);
        xlnt_assert_equals(general, xlnt::number_format::general());
        xlnt_assert_equals(general.id(), 0);
        xlnt_assert_equals(general.format_string(), "General");
    }

    void test_simple_format()
    {
        xlnt::number_format nf;

        nf.format_string("\"positive\"General;\"negative\"General");
        auto formatted = nf.format(3.14, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "positive3.14");
        formatted = nf.format(-3.14, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "negative3.14");

        nf.format_string("\"any\"General");
        formatted = nf.format(-3.14, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "-any3.14");

        nf.format_string("\"positive\"General;\"negative\"General;\"zero\"General");
        formatted = nf.format(3.14, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "positive3.14");
        formatted = nf.format(-3.14, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "negative3.14");
        formatted = nf.format(0, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "zero0");
    }

    void test_upper_case_date()
    {
        auto date = xlnt::date(2016, 6, 18);
        auto date_number = date.to_number(xlnt::calendar::windows_1900);

        xlnt::number_format nf;
        nf.format_string("MM / DD / YYYY");
        auto formatted = nf.format(date_number, xlnt::calendar::windows_1900);

        xlnt_assert_equals(formatted, "06 / 18 / 2016");
    }

    void test_simple_date()
    {
        auto date = xlnt::date(2016, 6, 18);
        auto date_number = date.to_number(xlnt::calendar::windows_1900);

        xlnt::number_format nf = xlnt::number_format::date_ddmmyyyy();
        auto formatted = nf.format(date_number, xlnt::calendar::windows_1900);

        xlnt_assert_equals(formatted, "18/06/16");
    }

    void test_short_month()
    {
        auto date = xlnt::date(2016, 6, 18);
        auto date_number = date.to_number(xlnt::calendar::windows_1900);

        xlnt::number_format nf;
        nf.format_string("m");
        auto formatted = nf.format(date_number, xlnt::calendar::windows_1900);

        xlnt_assert_equals(formatted, "6");
    }

    void test_month_abbreviation()
    {
        auto date = xlnt::date(2016, 6, 18);
        auto date_number = date.to_number(xlnt::calendar::windows_1900);

        xlnt::number_format nf;
        nf.format_string("mmm");
        auto formatted = nf.format(date_number, xlnt::calendar::windows_1900);

        xlnt_assert_equals(formatted, "Jun");
    }

    void test_month_name()
    {
        auto date = xlnt::date(2016, 6, 18);
        auto date_number = date.to_number(xlnt::calendar::windows_1900);

        xlnt::number_format nf;
        nf.format_string("mmmm");
        auto formatted = nf.format(date_number, xlnt::calendar::windows_1900);

        xlnt_assert_equals(formatted, "June");
    }

    void test_short_day()
    {
        auto date = xlnt::date(2016, 6, 8);
        auto date_number = date.to_number(xlnt::calendar::windows_1900);

        xlnt::number_format nf;
        nf.format_string("d");
        auto formatted = nf.format(date_number, xlnt::calendar::windows_1900);

        xlnt_assert_equals(formatted, "8");
    }

    void test_long_day()
    {
        auto date = xlnt::date(2016, 6, 8);
        auto date_number = date.to_number(xlnt::calendar::windows_1900);

        xlnt::number_format nf;
        nf.format_string("dd");
        auto formatted = nf.format(date_number, xlnt::calendar::windows_1900);

        xlnt_assert_equals(formatted, "08");
    }

    void test_long_year()
    {
        auto date = xlnt::date(2016, 6, 18);
        auto date_number = date.to_number(xlnt::calendar::windows_1900);

        xlnt::number_format nf;
        nf.format_string("yyyy");
        auto formatted = nf.format(date_number, xlnt::calendar::windows_1900);

        xlnt_assert_equals(formatted, "2016");
    }

    void test_day_name()
    {
        auto date = xlnt::date(2016, 6, 18);
        auto date_number = date.to_number(xlnt::calendar::windows_1900);

        xlnt::number_format nf;
        nf.format_string("dddd");
        auto formatted = nf.format(date_number, xlnt::calendar::windows_1900);

        xlnt_assert_equals(formatted, "Sunday");
    }

    void test_day_abbreviation()
    {
        auto date = xlnt::date(2016, 6, 18);
        auto date_number = date.to_number(xlnt::calendar::windows_1900);

        xlnt::number_format nf;
        nf.format_string("ddd");
        auto formatted = nf.format(date_number, xlnt::calendar::windows_1900);

        xlnt_assert_equals(formatted, "Sun");
    }

    void test_month_letter()
    {
        auto date = xlnt::date(2016, 6, 18);
        auto date_number = date.to_number(xlnt::calendar::windows_1900);

        xlnt::number_format nf;
        nf.format_string("mmmmm");
        auto formatted = nf.format(date_number, xlnt::calendar::windows_1900);

        xlnt_assert_equals(formatted, "J");
    }

    void test_time_24_hour()
    {
        auto time = xlnt::time(20, 15, 10);
        auto time_number = time.to_number();

        xlnt::number_format nf = xlnt::number_format::date_time4();
        auto formatted = nf.format(time_number, xlnt::calendar::windows_1900);

        xlnt_assert_equals(formatted, "20:15:10");
    }

    void test_elapsed_minutes()
    {
        auto period = xlnt::timedelta(1, 2, 3, 4, 5);
        auto period_number = period.to_number();

        xlnt::number_format nf("[mm]:ss");
        auto formatted = nf.format(period_number, xlnt::calendar::windows_1900);

        xlnt_assert_equals(formatted, "1563:04");
    }

    void test_second_fractional_leading_zero()
    {
        auto time = xlnt::time(1, 2, 3, 400000);
        auto time_number = time.to_number();

        xlnt::number_format nf("ss.0");
        auto formatted = nf.format(time_number, xlnt::calendar::windows_1900);

        xlnt_assert_equals(formatted, "03.4");
    }

    void test_second_fractional()
    {
        auto time = xlnt::time(1, 2, 3, 400000);
        auto time_number = time.to_number();

        xlnt::number_format nf("s.0");
        auto formatted = nf.format(time_number, xlnt::calendar::windows_1900);

        xlnt_assert_equals(formatted, "3.4");
    }

    void test_elapsed_seconds()
    {
        auto period = xlnt::timedelta(1, 2, 3, 4, 5);
        auto period_number = period.to_number();

        xlnt::number_format nf("[ss]");
        auto formatted = nf.format(period_number, xlnt::calendar::windows_1900);

        xlnt_assert_equals(formatted, "93784");
    }

    void test_time_12_hour_am()
    {
        auto time = xlnt::time(8, 15, 10);
        auto time_number = time.to_number();

        xlnt::number_format nf = xlnt::number_format::date_time2();
        auto formatted = nf.format(time_number, xlnt::calendar::windows_1900);

        xlnt_assert_equals(formatted, "8:15:10 AM");
    }

    void test_time_12_hour_pm()
    {
        auto time = xlnt::time(20, 15, 10);
        auto time_number = time.to_number();

        xlnt::number_format nf = xlnt::number_format::date_time2();
        auto formatted = nf.format(time_number, xlnt::calendar::windows_1900);

        xlnt_assert_equals(formatted, "8:15:10 PM");
    }

    void test_long_hour_12_hour()
    {
        auto time = xlnt::time(20, 15, 10);
        auto time_number = time.to_number();

        xlnt::number_format nf;
        nf.format_string("hh AM/PM");
        auto formatted = nf.format(time_number, xlnt::calendar::windows_1900);

        xlnt_assert_equals(formatted, "08 PM");
    }

    void test_long_hour_12_hour_ap()
    {
        auto time1 = xlnt::time(20, 15, 10);
        auto time1_number = time1.to_number();

        auto time2 = xlnt::time(8, 15, 10);
        auto time2_number = time2.to_number();

        xlnt::number_format nf("hh A/P");

        auto formatted = nf.format(time1_number, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "08 P");

        formatted = nf.format(time2_number, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "08 A");
    }

    void test_long_hour_24_hour()
    {
        auto time = xlnt::time(20, 15, 10);
        auto time_number = time.to_number();

        xlnt::number_format nf;
        nf.format_string("hh");
        auto formatted = nf.format(time_number, xlnt::calendar::windows_1900);

        xlnt_assert_equals(formatted, "20");
    }

    void test_short_minute()
    {
        auto time = xlnt::time(20, 5, 10);
        auto time_number = time.to_number();

        xlnt::number_format nf;
        nf.format_string("h:m");
        auto formatted = nf.format(time_number, xlnt::calendar::windows_1900);

        xlnt_assert_equals(formatted, "20:5");
    }

    void test_long_minute()
    {
        auto time = xlnt::time(20, 5, 10);
        auto time_number = time.to_number();

        xlnt::number_format nf;
        nf.format_string("h:mm");
        auto formatted = nf.format(time_number, xlnt::calendar::windows_1900);

        xlnt_assert_equals(formatted, "20:05");
    }

    void test_short_second()
    {
        auto time = xlnt::time(20, 15, 1);
        auto time_number = time.to_number();

        xlnt::number_format nf;
        nf.format_string("h:m:s");
        auto formatted = nf.format(time_number, xlnt::calendar::windows_1900);

        xlnt_assert_equals(formatted, "20:15:1");
    }

    void test_long_second()
    {
        auto time = xlnt::time(20, 15, 1);
        auto time_number = time.to_number();

        xlnt::number_format nf;
        nf.format_string("h:m:ss");
        auto formatted = nf.format(time_number, xlnt::calendar::windows_1900);

        xlnt_assert_equals(formatted, "20:15:01");
    }

    void test_trailing_space()
    {
        auto time = xlnt::time(20, 15, 1);
        auto time_number = time.to_number();

        xlnt::number_format nf;
        nf.format_string("h:m:ss ");
        auto formatted = nf.format(time_number, xlnt::calendar::windows_1900);

        xlnt_assert_equals(formatted, "20:15:01 ");
    }

    void test_text_section_string()
    {
        xlnt::number_format nf;
        nf.format_string("General;General;General;[Green]\"a\"@\"b\"");

        auto formatted = nf.format("text");

        xlnt_assert_equals(formatted, "atextb");
    }

    void test_text_section_no_string()
    {
        xlnt::number_format nf;
        nf.format_string("General;General;General;[Green]m\"ab\"");

        auto formatted = nf.format("text");

        xlnt_assert_equals(formatted, "ab");
    }

    void test_text_section_no_text()
    {
        xlnt::number_format nf;
        nf.format_string("General;General;General;[Green]m");

        auto formatted = nf.format("text");

        xlnt_assert_equals(formatted, "text");
    }

    void test_conditional_format()
    {
        xlnt::number_format nf;

        nf.format_string("[>5]General\"first\";[>3]\"second\"General;\"third\"General");
        auto formatted = nf.format(6, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "6first");
        formatted = nf.format(4, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "second4");
        formatted = nf.format(5, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "second5");
        formatted = nf.format(3, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "third3");
        formatted = nf.format(2, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "third2");

        nf.format_string("[>=5]\"first\"General;[>=3]\"second\"General;\"third\"General");
        formatted = nf.format(5, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "first5");
        formatted = nf.format(6, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "first6");
        formatted = nf.format(4, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "second4");
        formatted = nf.format(3, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "second3");
        formatted = nf.format(2, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "third2");

        nf.format_string("[>=5]\"first\"General");
        formatted = nf.format(4, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "###########");

        nf.format_string("[>=5]\"first\"General;[>=4]\"second\"General");
        formatted = nf.format(3, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "###########");

        nf.format_string("[<1]\"first\"General;[<5]\"second\"General;\"third\"General");
        formatted = nf.format(0, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "first0");
        formatted = nf.format(1, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "second1");
        formatted = nf.format(5, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "third5");
        formatted = nf.format(6, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "third6");

        nf.format_string("[<=1]\"first\"General;[<=5]\"second\"General;\"third\"General");
        formatted = nf.format(-1000, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "-first1000");
        formatted = nf.format(0, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "first0");
        formatted = nf.format(1, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "first1");
        formatted = nf.format(4, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "second4");
        formatted = nf.format(5, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "second5");
        formatted = nf.format(6, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "third6");
        formatted = nf.format(1000, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "third1000");

        nf.format_string("[=1]\"first\"General;[=2]\"second\"General;\"third\"General");
        formatted = nf.format(1, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "first1");
        formatted = nf.format(2, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "second2");
        formatted = nf.format(3, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "third3");
        formatted = nf.format(0, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "third0");

        nf.format_string("[<>1]\"first\"General;[<>2]\"second\"General");
        formatted = nf.format(2, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "first2");
        formatted = nf.format(1, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "second1");
    }

    void test_space()
    {
        xlnt::number_format nf;
        nf.format_string("_(General_)");
        auto formatted = nf.format(6, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, " 6 ");
    }

    void test_fill()
    {
        xlnt::number_format nf;
        nf.format_string("*-General");
        auto formatted = nf.format(6, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "----------6");

        nf.format_string("General*-");
        formatted = nf.format(6, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "6----------");

        nf.format_string("\\a*-\\b");
        formatted = nf.format(6, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "a---------b");
    }

    void test_placeholders_zero()
    {
        xlnt::number_format nf;
        nf.format_string("00");
        auto formatted = nf.format(6, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "06");

        nf.format_string("00");
        formatted = nf.format(63, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "63");
    }

    void test_placeholders_space()
    {
        xlnt::number_format nf;
        nf.format_string("?0");
        auto formatted = nf.format(6, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, " 6");

        nf.format_string("?0");
        formatted = nf.format(63, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "63");

        nf.format_string("?0");
        formatted = nf.format(637, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "637");

        nf.format_string("0.?");
        formatted = nf.format(6, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "6. ");

        nf.format_string("0.0?");
        formatted = nf.format(6.3, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "6.3 ");
        formatted = nf.format(6.34, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "6.34");
    }

    void test_scientific()
    {
        xlnt::number_format nf;
        nf.format_string("0.0E-0");
        auto formatted = nf.format(6.1, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "6.1E0");
    }

    void test_locale_currency()
    {
        xlnt::number_format nf;

        nf.format_string("[$€-407]#,##0.00");
        auto formatted = nf.format(-45000.1, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "-€45,000.10");

        nf.format_string("[$$-1009]#,##0.00");
        formatted = nf.format(-45000.1, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "-$45,000.10");
    }

    void test_bad_country()
    {
        xlnt::number_format nf;

        nf.format_string("[$-]#,##0.00");
        xlnt_assert_throws(nf.format(1.2, xlnt::calendar::windows_1900), std::runtime_error);

        nf.format_string("[$-G]#,##0.00");
        xlnt_assert_throws(nf.format(1.2, xlnt::calendar::windows_1900), std::runtime_error);

        nf.format_string("[$-4002]#,##0.00");
        xlnt_assert_throws(nf.format(1.2, xlnt::calendar::windows_1900), std::runtime_error);

        nf.format_string("[-4001]#,##0.00");
        xlnt_assert_throws(nf.format(1.2, xlnt::calendar::windows_1900), std::runtime_error);
    }

    void test_duplicate_bracket_sections()
    {
        xlnt::number_format nf;

        nf.format_string("[Red][Green]#,##0.00");
        xlnt_assert_throws(nf.format(1.2, xlnt::calendar::windows_1900), std::runtime_error);

        nf.format_string("[$-403][$-4001]#,##0.00");
        xlnt_assert_throws(nf.format(1.2, xlnt::calendar::windows_1900), std::runtime_error);

        nf.format_string("[>3][>4]#,##0.00");
        xlnt_assert_throws(nf.format(1.2, xlnt::calendar::windows_1900), std::runtime_error);
    }

    void test_escaped_quote_string()
    {
        xlnt::number_format nf;

        nf.format_string("\"\\\"\"General");
        auto formatted = nf.format(6, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "\"6");
    }

    void test_thousands_scale()
    {
        xlnt::number_format nf;

        nf.format_string("#,");
        auto formatted = nf.format(61234, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "61");
    }

    void test_colors()
    {
        xlnt::number_format nf;

        nf.format_string("[Black]#");
        auto formatted = nf.format(6, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "6");

        nf.format_string("[Black]#");
        formatted = nf.format(6, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "6");

        nf.format_string("[Blue]#");
        formatted = nf.format(6, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "6");

        nf.format_string("[Green]#");
        formatted = nf.format(6, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "6");

        nf.format_string("[Red]#");
        formatted = nf.format(6, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "6");

        nf.format_string("[Cyan]#");
        formatted = nf.format(6, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "6");

        nf.format_string("[Magenta]#");
        formatted = nf.format(6, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "6");

        nf.format_string("[Yellow]#");
        formatted = nf.format(6, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "6");

        nf.format_string("[White]#");
        formatted = nf.format(6, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "6");

        nf.format_string("[Color15]#");
        formatted = nf.format(6, xlnt::calendar::windows_1900);
        xlnt_assert_equals(formatted, "6");
    }

    void test_bad_format()
    {
        xlnt::number_format nf;

        nf.format_string("[=1]\"first\"General;[=2]\"second\"General;[=3]\"third\"General");
        xlnt_assert_throws(nf.format(1.2, xlnt::calendar::windows_1900), std::runtime_error);

        nf.format_string("\"first\"General;\"second\"General;\"third\"General;\"fourth\"General;\"fifth\"General");
        xlnt_assert_throws(nf.format(1.2, xlnt::calendar::windows_1900), std::runtime_error);

        nf.format_string("[");
        xlnt_assert_throws(nf.format(1.2, xlnt::calendar::windows_1900), std::runtime_error);

        nf.format_string("[]");
        xlnt_assert_throws(nf.format(1.2, xlnt::calendar::windows_1900), std::runtime_error);

        nf.format_string("[Redd]");
        xlnt_assert_throws(nf.format(1.2, xlnt::calendar::windows_1900), std::runtime_error);

        nf.format_string("[$1]#");
        xlnt_assert_throws(nf.format(1.2, xlnt::calendar::windows_1900), std::runtime_error);

        nf.format_string("Gee");
        xlnt_assert_throws(nf.format(1.2, xlnt::calendar::windows_1900), std::runtime_error);

        nf.format_string("!");
        xlnt_assert_throws(nf.format(1.2, xlnt::calendar::windows_1900), std::runtime_error);

        nf.format_string("A/");
        xlnt_assert_throws(nf.format(1.2, xlnt::calendar::windows_1900), std::runtime_error);
    }

    void format_and_test(const xlnt::number_format &nf, const std::array<std::string, 4> &expect)
    {
        double positive = 42503.1234;
        double negative = -1 * positive;
        double zero = 0;
        const std::string text = "text";
        xlnt::calendar calendar = xlnt::calendar::windows_1900;

        xlnt_assert_equals(nf.format(positive, calendar), expect[0]);
        xlnt_assert_equals(nf.format(negative, calendar), expect[1]);
        xlnt_assert_equals(nf.format(zero, calendar), expect[2]);
        xlnt_assert_equals(nf.format(text), expect[3]);
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
        format_and_test(xlnt::number_format::number_comma_separated1(), {{"42,503.12", "-42,503.12", "0.00", "text"}});
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
    void test_builtin_format_12()
    {
        format_and_test(xlnt::number_format::from_builtin_id(12), {{"42503 1/8", "-42503 1/8", "0", "text"}});
    }

    // # ??/??
    void test_builtin_format_13()
    {
        format_and_test(xlnt::number_format::from_builtin_id(13), {{"42503 10/81", "-42503 10/81", "0", "text"}});
    }

    // mm-dd-yy
    void test_builtin_format_14()
    {
        format_and_test(xlnt::number_format::date_xlsx14(), {{"05-13-16", "###########", "01-00-00", "text"}});
    }

    // d-mmm-yy
    void test_builtin_format_15()
    {
        format_and_test(xlnt::number_format::date_xlsx15(), {{"13-May-16", "###########", "0-Jan-00", "text"}});
    }

    // d-mmm
    void test_builtin_format_16()
    {
        format_and_test(xlnt::number_format::date_xlsx16(), {{"13-May", "###########", "0-Jan", "text"}});
    }

    // mmm-yy
    void test_builtin_format_17()
    {
        format_and_test(xlnt::number_format::date_xlsx17(), {{"May-16", "###########", "Jan-00", "text"}});
    }

    // h:mm AM/PM
    void test_builtin_format_18()
    {
        format_and_test(xlnt::number_format::date_time1(), {{"2:57 AM", "###########", "12:00 AM", "text"}});
    }

    // h:mm:ss AM/PM
    void test_builtin_format_19()
    {
        format_and_test(xlnt::number_format::date_time2(), {{"2:57:42 AM", "###########", "12:00:00 AM", "text"}});
    }

    // h:mm
    void test_builtin_format_20()
    {
        format_and_test(xlnt::number_format::date_time3(), {{"2:57", "###########", "0:00", "text"}});
    }

    // h:mm:ss
    void test_builtin_format_21()
    {
        format_and_test(xlnt::number_format::date_time4(), {{"2:57:42", "###########", "0:00:00", "text"}});
    }

    // m/d/yy h:mm
    void test_builtin_format_22()
    {
        format_and_test(xlnt::number_format::date_xlsx22(), {{"5/13/16 2:57", "###########", "1/0/00 0:00", "text"}});
    }

    // #,##0 ;(#,##0)
    void test_builtin_format_37()
    {
        format_and_test(xlnt::number_format::from_builtin_id(37), {{"42,503 ", "(42,503)", "0 ", "text"}});
    }

    // #,##0 ;[Red](#,##0)
    void test_builtin_format_38()
    {
        format_and_test(xlnt::number_format::from_builtin_id(38), {{"42,503 ", "(42,503)", "0 ", "text"}});
    }

    // #,##0.00;(#,##0.00)
    void test_builtin_format_39()
    {
        format_and_test(xlnt::number_format::from_builtin_id(39), {{"42,503.12", "(42,503.12)", "0.00", "text"}});
    }

    // #,##0.00;[Red](#,##0.00)
    void test_builtin_format_40()
    {
        format_and_test(xlnt::number_format::from_builtin_id(40), {{"42,503.12", "(42,503.12)", "0.00", "text"}});
    }

    // mm:ss
    void test_builtin_format_45()
    {
        format_and_test(xlnt::number_format::date_time5(), {{"57:42", "###########", "00:00", "text"}});
    }

    // [h]:mm:ss
    void test_builtin_format_46()
    {
        format_and_test(xlnt::number_format::from_builtin_id(46), {{"1020074:57:42", "###########", "0:00:00", "text"}});
    }

    // mmss.0
    void test_builtin_format_47()
    {
        format_and_test(xlnt::number_format::from_builtin_id(47), {{"5741.8", "###########", "0000.0", "text"}});
    }

    // ##0.0E+0
    void test_builtin_format_48()
    {
        format_and_test(xlnt::number_format::from_builtin_id(48), {{"42.5E+3", "-42.5E+3", "000.0E+0", "text"}});
    }

    // @
    void test_builtin_format_49()
    {
        format_and_test(xlnt::number_format::text(), {{"42503.1234", "-42503.1234", "0", "text"}});
    }

    // yy-mm-dd
    void test_builtin_format_date_yyyymmdd()
    {
        format_and_test(xlnt::number_format::date_yymmdd(), {{"16-05-13", "###########", "00-01-00", "text"}});
    }

    // d/m/y
    void test_builtin_format_date_dmyslash()
    {
        format_and_test(xlnt::number_format::date_dmyslash(), {{"13/5/16", "###########", "0/1/00", "text"}});
    }

    // d-m-y
    void test_builtin_format_date_dmyminus()
    {
        format_and_test(xlnt::number_format::date_dmyminus(), {{"13-5-16", "###########", "0-1-00", "text"}});
    }

    // d-m
    void test_builtin_format_date_dmminus()
    {
        format_and_test(xlnt::number_format::date_dmminus(), {{"13-5", "###########", "0-1", "text"}});
    }

    // m-yy
    void test_builtin_format_date_myminus()
    {
        format_and_test(xlnt::number_format::date_myminus(), {{"5-16", "###########", "1-00", "text"}});
    }
};
static number_format_test_suite x;