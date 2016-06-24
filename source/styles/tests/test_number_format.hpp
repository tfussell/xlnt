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

        nf.set_format_string("\"positive\"General;\"negative\"General");
        formatted = nf.format(-3.14, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "-negative3.14");
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

        nf.set_format_string("[$€-407]-#,##0.00");
        auto formatted = nf.format(1.2, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "-€1,20");

        nf.set_format_string("[$$-1009]#,##0.00");
        formatted = nf.format(1.2, xlnt::calendar::windows_1900);
        TS_ASSERT_EQUALS(formatted, "CA$1,20");
        
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

        nf.set_format_string("");
        TS_ASSERT_THROWS(nf.format(1.2, xlnt::calendar::windows_1900), std::runtime_error);

        nf.set_format_string(";;;");
        TS_ASSERT_THROWS(nf.format(1.2, xlnt::calendar::windows_1900), std::runtime_error);

        nf.set_format_string("[=1]\"first\"General;\"second\"General;\"third\"General");
        TS_ASSERT_THROWS(nf.format(1.2, xlnt::calendar::windows_1900), std::runtime_error);

        nf.set_format_string("[=1]\"first\"General;[=2]\"second\"General;[=3]\"third\"General");
        TS_ASSERT_THROWS(nf.format(1.2, xlnt::calendar::windows_1900), std::runtime_error);

        nf.set_format_string("[=1]\"first\"General;[=2]\"second\"General;\"third\"General;\"fourth\"General");
        TS_ASSERT_THROWS(nf.format(1.2, xlnt::calendar::windows_1900), std::runtime_error);

        nf.set_format_string("\"first\"General;\"second\"General;\"third\"General;\"fourth\"General;\"fifth\"General");
        TS_ASSERT_THROWS(nf.format(1.2, xlnt::calendar::windows_1900), std::runtime_error);
    }

    void test_builtin_formats()
    {
        TS_ASSERT_EQUALS(xlnt::number_format::text().format("a"), "a");
        TS_ASSERT_EQUALS(xlnt::number_format::number().format(1, xlnt::calendar::windows_1900), "1");
        TS_ASSERT_EQUALS(xlnt::number_format::number_comma_separated1().format(1, xlnt::calendar::windows_1900), "1,00");

/*
        auto datetime = xlnt::datetime(2016, 6, 24, 0, 45, 58);
        auto datetime_number = datetime.to_number(xlnt::calendar::windows_1900);

        TS_ASSERT_EQUALS(xlnt::number_format::percentage().format(datetime_number, xlnt::calendar::windows_1900), "1");
        TS_ASSERT_EQUALS(xlnt::number_format::percentage_00().format(datetime_number, xlnt::calendar::windows_1900), "1");
        TS_ASSERT_EQUALS(xlnt::number_format::date_yyyymmdd2().format(datetime_number, xlnt::calendar::windows_1900), "1");
        TS_ASSERT_EQUALS(xlnt::number_format::date_yyyymmdd().format(datetime_number, xlnt::calendar::windows_1900), "1");
        TS_ASSERT_EQUALS(xlnt::number_format::date_ddmmyyyy().format(datetime_number, xlnt::calendar::windows_1900), "1");
        TS_ASSERT_EQUALS(xlnt::number_format::date_dmyslash().format(datetime_number, xlnt::calendar::windows_1900), "1");
        TS_ASSERT_EQUALS(xlnt::number_format::date_dmyminus().format(datetime_number, xlnt::calendar::windows_1900), "1");
        TS_ASSERT_EQUALS(xlnt::number_format::date_dmminus().format(datetime_number, xlnt::calendar::windows_1900), "1");
        TS_ASSERT_EQUALS(xlnt::number_format::date_myminus().format(datetime_number, xlnt::calendar::windows_1900), "1");
        TS_ASSERT_EQUALS(xlnt::number_format::date_xlsx14().format(datetime_number, xlnt::calendar::windows_1900), "1");
        TS_ASSERT_EQUALS(xlnt::number_format::date_xlsx15().format(datetime_number, xlnt::calendar::windows_1900), "1");
        TS_ASSERT_EQUALS(xlnt::number_format::date_xlsx16().format(datetime_number, xlnt::calendar::windows_1900), "1");
        TS_ASSERT_EQUALS(xlnt::number_format::date_xlsx17().format(datetime_number, xlnt::calendar::windows_1900), "1");
        TS_ASSERT_EQUALS(xlnt::number_format::date_xlsx22().format(datetime_number, xlnt::calendar::windows_1900), "1");
        TS_ASSERT_EQUALS(xlnt::number_format::date_datetime().format(datetime_number, xlnt::calendar::windows_1900), "1");
        TS_ASSERT_EQUALS(xlnt::number_format::date_time1().format(datetime_number, xlnt::calendar::windows_1900), "1");
        TS_ASSERT_EQUALS(xlnt::number_format::date_time2().format(datetime_number, xlnt::calendar::windows_1900), "1");
        TS_ASSERT_EQUALS(xlnt::number_format::date_time3().format(datetime_number, xlnt::calendar::windows_1900), "1");
        TS_ASSERT_EQUALS(xlnt::number_format::date_time4().format(datetime_number, xlnt::calendar::windows_1900), "1");
        TS_ASSERT_EQUALS(xlnt::number_format::date_time5().format(datetime_number, xlnt::calendar::windows_1900), "1");
        TS_ASSERT_EQUALS(xlnt::number_format::date_time6().format(datetime_number, xlnt::calendar::windows_1900), "1");
        TS_ASSERT_EQUALS(xlnt::number_format::date_time7().format(datetime_number, xlnt::calendar::windows_1900), "1");
        TS_ASSERT_EQUALS(xlnt::number_format::date_time8().format(datetime_number, xlnt::calendar::windows_1900), "1");
        TS_ASSERT_EQUALS(xlnt::number_format::date_timedelta().format(datetime_number, xlnt::calendar::windows_1900), "1");
        TS_ASSERT_EQUALS(xlnt::number_format::date_yyyymmddslash().format(datetime_number, xlnt::calendar::windows_1900), "1");

        TS_ASSERT_EQUALS(xlnt::number_format::currency_usd_simple().format(1.23, xlnt::calendar::windows_1900), "1");
        TS_ASSERT_EQUALS(xlnt::number_format::currency_usd().format(1.23, xlnt::calendar::windows_1900), "1");
        TS_ASSERT_EQUALS(xlnt::number_format::currency_eur_simple().format(1.23, xlnt::calendar::windows_1900), "1");
*/
    }
};
