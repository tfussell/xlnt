#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/xlnt.h>

class NumberFormatTestSuite : public CxxTest::TestSuite
{
public:
    NumberFormatTestSuite()
    {

    }

    void setup_class(int cls)
    {
        //cls.workbook = Workbook()
        //    cls.worksheet = Worksheet(cls.workbook, "Test")
        //    cls.sd = SharedDate()
    }

    void test_convert_date_to_julian()
    {
        //TS_ASSERT_EQUALS(40167, sd.to_julian(2009, 12, 20))
    }

    void test_convert_date_from_julian()
    {
    }

    void test_date_equal(int julian, int datetime)
    {
        //TS_ASSERT_EQUALS(sd.from_julian(julian), datetime);

        //date_pairs = (
        //    (40167, datetime(2009, 12, 20)),
        //    (21980, datetime(1960, 3, 5)),
        //    );

        //for count, dt in date_pairs
        //{
        //    yield test_date_equal, count, dt;
        //}
    }

    void test_convert_datetime_to_julian()
    {
        //TS_ASSERT_EQUALS(40167, sd.datetime_to_julian(datetime(2009, 12, 20)))
        //    TS_ASSERT_EQUALS(40196.5939815, sd.datetime_to_julian(datetime(2010, 1, 18, 14, 15, 20, 1600)))
    }

    void test_insert_float()
    {
        //worksheet.cell("A1").value = 3.14
        //    TS_ASSERT_EQUALS(Cell.TYPE_NUMERIC, worksheet.cell("A1")._data_type)
    }

    void test_insert_percentage()
    {
        //worksheet.cell("A1").value = "3.14%"
        //    TS_ASSERT_EQUALS(Cell.TYPE_NUMERIC, worksheet.cell("A1")._data_type)
        //    assert_almost_equal(0.0314, worksheet.cell("A1").value)
    }

    void test_insert_datetime()
    {
        //worksheet.cell("A1").value = date.today()
        //    TS_ASSERT_EQUALS(Cell.TYPE_NUMERIC, worksheet.cell("A1")._data_type)
    }

    void test_insert_date()
    {
        //worksheet.cell("A1").value = datetime.now()
        //    TS_ASSERT_EQUALS(Cell.TYPE_NUMERIC, worksheet.cell("A1")._data_type)
    }

    void test_internal_date()
    {
        //dt = datetime(2010, 7, 13, 6, 37, 41)
        //    worksheet.cell("A3").value = dt
        //    TS_ASSERT_EQUALS(40372.27616898148, worksheet.cell("A3")._value)
    }

    void test_datetime_interpretation()
    {
        //dt = datetime(2010, 7, 13, 6, 37, 41)
        //    worksheet.cell("A3").value = dt
        //    TS_ASSERT_EQUALS(dt, worksheet.cell("A3").value)
    }

    void test_date_interpretation()
    {
        //dt = date(2010, 7, 13)
        //    worksheet.cell("A3").value = dt
        //    TS_ASSERT_EQUALS(datetime(2010, 7, 13, 0, 0), worksheet.cell("A3").value)
    }

    void test_number_format_style()
    {
        //worksheet.cell("A1").value = "12.6%"
        //    TS_ASSERT_EQUALS(NumberFormat.FORMAT_PERCENTAGE, \
        //    worksheet.cell("A1").style.number_format.format_code)
    }

    void test_date_format_on_non_date()
    {
        //cell = worksheet.cell("A1");
    }

    void check_date_pair(int count, const std::string &date_string)
    {
        //cell.value = strptime(date_string, "%Y-%m-%d");
        //TS_ASSERT_EQUALS(count, cell._value);

        //date_pairs = (
        //    (15, "1900-01-15"),
        //    (59, "1900-02-28"),
        //    (61, "1900-03-01"),
        //    (367, "1901-01-01"),
        //    (2958465, "9999-12-31"), );
        //for count, date_string in date_pairs
        //{
        //    yield check_date_pair, count, date_string;
        //}
    }

    void test_1900_leap_year()
    {
        //assert_raises(ValueError, sd.from_julian, 60)
        //    assert_raises(ValueError, sd.to_julian, 1900, 2, 29)
    }

    void test_bad_date()
    {
        //void check_bad_date(year, month, day)
        //{
        //    assert_raises(ValueError, sd.to_julian, year, month, day)
        //}

        //bad_dates = ((1776, 7, 4), (1899, 12, 31), )
        //    for year, month, day in bad_dates
        //    {
        //    yield check_bad_date, year, month, day
        //    }
    }

    void test_bad_julian_date()
    {
        //assert_raises(ValueError, sd.from_julian, -1)
    }

    void test_mac_date()
    {
    //    sd.excel_base_date = CALENDAR_MAC_1904

    //        datetuple = (2011, 10, 31)

    //        dt = date(datetuple[0], datetuple[1], datetuple[2])
    //        julian = sd.to_julian(datetuple[0], datetuple[1], datetuple[2])
    //        reverse = sd.from_julian(julian).date()
    //        TS_ASSERT_EQUALS(dt, reverse)
    //        sd.excel_base_date = CALENDAR_WINDOWS_1900
    }
};
