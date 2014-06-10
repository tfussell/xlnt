/* Generated file, do not edit */

#ifndef CXXTEST_RUNNING
#define CXXTEST_RUNNING
#endif

#define _CXXTEST_HAVE_STD
#define _CXXTEST_HAVE_EH
#include <cxxtest/TestListener.h>
#include <cxxtest/TestTracker.h>
#include <cxxtest/TestRunner.h>
#include <cxxtest/RealDescriptions.h>
#include <cxxtest/TestMain.h>
#include <cxxtest/ErrorPrinter.h>

int main( int argc, char *argv[] ) {
 int status;
    CxxTest::ErrorPrinter tmp;
    CxxTest::RealWorldDescription::_worldName = "cxxtest";
    status = CxxTest::Main< CxxTest::ErrorPrinter >( tmp, argc, argv );
    return status;
}
bool suite_test_cell_init = false;
#include "/Users/thomas/Development/xlnt/tests/test_cell.hpp"

static test_cell suite_test_cell;

static CxxTest::List Tests_test_cell = { 0, 0 };
CxxTest::StaticSuiteDescription suiteDescription_test_cell( "../../tests/test_cell.hpp", 9, "test_cell", suite_test_cell, Tests_test_cell );

static class TestDescription_suite_test_cell_test_coordinates : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_cell_test_coordinates() : CxxTest::RealTestDescription( Tests_test_cell, suiteDescription_test_cell, 12, "test_coordinates" ) {}
 void runTest() { suite_test_cell.test_coordinates(); }
} testDescription_suite_test_cell_test_coordinates;

static class TestDescription_suite_test_cell_test_invalid_coordinate : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_cell_test_invalid_coordinate() : CxxTest::RealTestDescription( Tests_test_cell, suiteDescription_test_cell, 19, "test_invalid_coordinate" ) {}
 void runTest() { suite_test_cell.test_invalid_coordinate(); }
} testDescription_suite_test_cell_test_invalid_coordinate;

static class TestDescription_suite_test_cell_test_zero_row : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_cell_test_zero_row() : CxxTest::RealTestDescription( Tests_test_cell, suiteDescription_test_cell, 25, "test_zero_row" ) {}
 void runTest() { suite_test_cell.test_zero_row(); }
} testDescription_suite_test_cell_test_zero_row;

static class TestDescription_suite_test_cell_test_absolute : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_cell_test_absolute() : CxxTest::RealTestDescription( Tests_test_cell, suiteDescription_test_cell, 31, "test_absolute" ) {}
 void runTest() { suite_test_cell.test_absolute(); }
} testDescription_suite_test_cell_test_absolute;

static class TestDescription_suite_test_cell_test_absolute_multiple : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_cell_test_absolute_multiple() : CxxTest::RealTestDescription( Tests_test_cell, suiteDescription_test_cell, 36, "test_absolute_multiple" ) {}
 void runTest() { suite_test_cell.test_absolute_multiple(); }
} testDescription_suite_test_cell_test_absolute_multiple;

static class TestDescription_suite_test_cell_test_column_index : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_cell_test_column_index() : CxxTest::RealTestDescription( Tests_test_cell, suiteDescription_test_cell, 41, "test_column_index" ) {}
 void runTest() { suite_test_cell.test_column_index(); }
} testDescription_suite_test_cell_test_column_index;

static class TestDescription_suite_test_cell_test_bad_column_index : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_cell_test_bad_column_index() : CxxTest::RealTestDescription( Tests_test_cell, suiteDescription_test_cell, 71, "test_bad_column_index" ) {}
 void runTest() { suite_test_cell.test_bad_column_index(); }
} testDescription_suite_test_cell_test_bad_column_index;

static class TestDescription_suite_test_cell_test_column_letter_boundries : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_cell_test_column_letter_boundries() : CxxTest::RealTestDescription( Tests_test_cell, suiteDescription_test_cell, 79, "test_column_letter_boundries" ) {}
 void runTest() { suite_test_cell.test_column_letter_boundries(); }
} testDescription_suite_test_cell_test_column_letter_boundries;

static class TestDescription_suite_test_cell_test_column_letter : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_cell_test_column_letter() : CxxTest::RealTestDescription( Tests_test_cell, suiteDescription_test_cell, 88, "test_column_letter" ) {}
 void runTest() { suite_test_cell.test_column_letter(); }
} testDescription_suite_test_cell_test_column_letter;

static class TestDescription_suite_test_cell_test_initial_value : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_cell_test_initial_value() : CxxTest::RealTestDescription( Tests_test_cell, suiteDescription_test_cell, 98, "test_initial_value" ) {}
 void runTest() { suite_test_cell.test_initial_value(); }
} testDescription_suite_test_cell_test_initial_value;

static class TestDescription_suite_test_cell_test_1st : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_cell_test_1st() : CxxTest::RealTestDescription( Tests_test_cell, suiteDescription_test_cell, 106, "test_1st" ) {}
 void runTest() { suite_test_cell.test_1st(); }
} testDescription_suite_test_cell_test_1st;

static class TestDescription_suite_test_cell_test_null : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_cell_test_null() : CxxTest::RealTestDescription( Tests_test_cell, suiteDescription_test_cell, 114, "test_null" ) {}
 void runTest() { suite_test_cell.test_null(); }
} testDescription_suite_test_cell_test_null;

static class TestDescription_suite_test_cell_test_numeric : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_cell_test_numeric() : CxxTest::RealTestDescription( Tests_test_cell, suiteDescription_test_cell, 123, "test_numeric" ) {}
 void runTest() { suite_test_cell.test_numeric(); }
} testDescription_suite_test_cell_test_numeric;

static class TestDescription_suite_test_cell_test_string : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_cell_test_string() : CxxTest::RealTestDescription( Tests_test_cell, suiteDescription_test_cell, 165, "test_string" ) {}
 void runTest() { suite_test_cell.test_string(); }
} testDescription_suite_test_cell_test_string;

static class TestDescription_suite_test_cell_test_single_dot : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_cell_test_single_dot() : CxxTest::RealTestDescription( Tests_test_cell, suiteDescription_test_cell, 174, "test_single_dot" ) {}
 void runTest() { suite_test_cell.test_single_dot(); }
} testDescription_suite_test_cell_test_single_dot;

static class TestDescription_suite_test_cell_test_formula : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_cell_test_formula() : CxxTest::RealTestDescription( Tests_test_cell, suiteDescription_test_cell, 182, "test_formula" ) {}
 void runTest() { suite_test_cell.test_formula(); }
} testDescription_suite_test_cell_test_formula;

static class TestDescription_suite_test_cell_test_boolean : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_cell_test_boolean() : CxxTest::RealTestDescription( Tests_test_cell, suiteDescription_test_cell, 192, "test_boolean" ) {}
 void runTest() { suite_test_cell.test_boolean(); }
} testDescription_suite_test_cell_test_boolean;

static class TestDescription_suite_test_cell_test_leading_zero : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_cell_test_leading_zero() : CxxTest::RealTestDescription( Tests_test_cell, suiteDescription_test_cell, 202, "test_leading_zero" ) {}
 void runTest() { suite_test_cell.test_leading_zero(); }
} testDescription_suite_test_cell_test_leading_zero;

static class TestDescription_suite_test_cell_test_error_codes : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_cell_test_error_codes() : CxxTest::RealTestDescription( Tests_test_cell, suiteDescription_test_cell, 210, "test_error_codes" ) {}
 void runTest() { suite_test_cell.test_error_codes(); }
} testDescription_suite_test_cell_test_error_codes;

static class TestDescription_suite_test_cell_test_insert_float : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_cell_test_insert_float() : CxxTest::RealTestDescription( Tests_test_cell, suiteDescription_test_cell, 222, "test_insert_float" ) {}
 void runTest() { suite_test_cell.test_insert_float(); }
} testDescription_suite_test_cell_test_insert_float;

static class TestDescription_suite_test_cell_test_insert_percentage : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_cell_test_insert_percentage() : CxxTest::RealTestDescription( Tests_test_cell, suiteDescription_test_cell, 230, "test_insert_percentage" ) {}
 void runTest() { suite_test_cell.test_insert_percentage(); }
} testDescription_suite_test_cell_test_insert_percentage;

static class TestDescription_suite_test_cell_test_insert_datetime : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_cell_test_insert_datetime() : CxxTest::RealTestDescription( Tests_test_cell, suiteDescription_test_cell, 238, "test_insert_datetime" ) {}
 void runTest() { suite_test_cell.test_insert_datetime(); }
} testDescription_suite_test_cell_test_insert_datetime;

static class TestDescription_suite_test_cell_test_insert_date : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_cell_test_insert_date() : CxxTest::RealTestDescription( Tests_test_cell, suiteDescription_test_cell, 246, "test_insert_date" ) {}
 void runTest() { suite_test_cell.test_insert_date(); }
} testDescription_suite_test_cell_test_insert_date;

static class TestDescription_suite_test_cell_test_internal_date : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_cell_test_internal_date() : CxxTest::RealTestDescription( Tests_test_cell, suiteDescription_test_cell, 254, "test_internal_date" ) {}
 void runTest() { suite_test_cell.test_internal_date(); }
} testDescription_suite_test_cell_test_internal_date;

static class TestDescription_suite_test_cell_test_datetime_interpretation : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_cell_test_datetime_interpretation() : CxxTest::RealTestDescription( Tests_test_cell, suiteDescription_test_cell, 263, "test_datetime_interpretation" ) {}
 void runTest() { suite_test_cell.test_datetime_interpretation(); }
} testDescription_suite_test_cell_test_datetime_interpretation;

static class TestDescription_suite_test_cell_test_date_interpretation : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_cell_test_date_interpretation() : CxxTest::RealTestDescription( Tests_test_cell, suiteDescription_test_cell, 273, "test_date_interpretation" ) {}
 void runTest() { suite_test_cell.test_date_interpretation(); }
} testDescription_suite_test_cell_test_date_interpretation;

static class TestDescription_suite_test_cell_test_number_format_style : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_cell_test_number_format_style() : CxxTest::RealTestDescription( Tests_test_cell, suiteDescription_test_cell, 283, "test_number_format_style" ) {}
 void runTest() { suite_test_cell.test_number_format_style(); }
} testDescription_suite_test_cell_test_number_format_style;

static class TestDescription_suite_test_cell_test_data_type_check : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_cell_test_data_type_check() : CxxTest::RealTestDescription( Tests_test_cell, suiteDescription_test_cell, 291, "test_data_type_check" ) {}
 void runTest() { suite_test_cell.test_data_type_check(); }
} testDescription_suite_test_cell_test_data_type_check;

static class TestDescription_suite_test_cell_test_set_bad_type : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_cell_test_set_bad_type() : CxxTest::RealTestDescription( Tests_test_cell, suiteDescription_test_cell, 308, "test_set_bad_type" ) {}
 void runTest() { suite_test_cell.test_set_bad_type(); }
} testDescription_suite_test_cell_test_set_bad_type;

static class TestDescription_suite_test_cell_test_time : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_cell_test_time() : CxxTest::RealTestDescription( Tests_test_cell, suiteDescription_test_cell, 326, "test_time" ) {}
 void runTest() { suite_test_cell.test_time(); }
} testDescription_suite_test_cell_test_time;

static class TestDescription_suite_test_cell_test_date_format_on_non_date : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_cell_test_date_format_on_non_date() : CxxTest::RealTestDescription( Tests_test_cell, suiteDescription_test_cell, 340, "test_date_format_on_non_date" ) {}
 void runTest() { suite_test_cell.test_date_format_on_non_date(); }
} testDescription_suite_test_cell_test_date_format_on_non_date;

static class TestDescription_suite_test_cell_test_set_get_date : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_cell_test_set_get_date() : CxxTest::RealTestDescription( Tests_test_cell, suiteDescription_test_cell, 350, "test_set_get_date" ) {}
 void runTest() { suite_test_cell.test_set_get_date(); }
} testDescription_suite_test_cell_test_set_get_date;

static class TestDescription_suite_test_cell_test_repr : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_cell_test_repr() : CxxTest::RealTestDescription( Tests_test_cell, suiteDescription_test_cell, 361, "test_repr" ) {}
 void runTest() { suite_test_cell.test_repr(); }
} testDescription_suite_test_cell_test_repr;

static class TestDescription_suite_test_cell_test_is_date : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_cell_test_is_date() : CxxTest::RealTestDescription( Tests_test_cell, suiteDescription_test_cell, 369, "test_is_date" ) {}
 void runTest() { suite_test_cell.test_is_date(); }
} testDescription_suite_test_cell_test_is_date;

static class TestDescription_suite_test_cell_test_is_not_date_color_format : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_cell_test_is_not_date_color_format() : CxxTest::RealTestDescription( Tests_test_cell, suiteDescription_test_cell, 383, "test_is_not_date_color_format" ) {}
 void runTest() { suite_test_cell.test_is_not_date_color_format(); }
} testDescription_suite_test_cell_test_is_not_date_color_format;

#include "/Users/thomas/Development/xlnt/tests/test_chart.hpp"

static test_chart suite_test_chart;

static CxxTest::List Tests_test_chart = { 0, 0 };
CxxTest::StaticSuiteDescription suiteDescription_test_chart( "../../tests/test_chart.hpp", 8, "test_chart", suite_test_chart, Tests_test_chart );

static class TestDescription_suite_test_chart_test_write_title : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_chart_test_write_title() : CxxTest::RealTestDescription( Tests_test_chart, suiteDescription_test_chart, 31, "test_write_title" ) {}
 void runTest() { suite_test_chart.test_write_title(); }
} testDescription_suite_test_chart_test_write_title;

static class TestDescription_suite_test_chart_test_write_xaxis : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_chart_test_write_xaxis() : CxxTest::RealTestDescription( Tests_test_chart, suiteDescription_test_chart, 39, "test_write_xaxis" ) {}
 void runTest() { suite_test_chart.test_write_xaxis(); }
} testDescription_suite_test_chart_test_write_xaxis;

static class TestDescription_suite_test_chart_test_write_yaxis : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_chart_test_write_yaxis() : CxxTest::RealTestDescription( Tests_test_chart, suiteDescription_test_chart, 47, "test_write_yaxis" ) {}
 void runTest() { suite_test_chart.test_write_yaxis(); }
} testDescription_suite_test_chart_test_write_yaxis;

static class TestDescription_suite_test_chart_test_write_series : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_chart_test_write_series() : CxxTest::RealTestDescription( Tests_test_chart, suiteDescription_test_chart, 55, "test_write_series" ) {}
 void runTest() { suite_test_chart.test_write_series(); }
} testDescription_suite_test_chart_test_write_series;

static class TestDescription_suite_test_chart_test_write_legend : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_chart_test_write_legend() : CxxTest::RealTestDescription( Tests_test_chart, suiteDescription_test_chart, 61, "test_write_legend" ) {}
 void runTest() { suite_test_chart.test_write_legend(); }
} testDescription_suite_test_chart_test_write_legend;

static class TestDescription_suite_test_chart_test_no_write_legend : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_chart_test_no_write_legend() : CxxTest::RealTestDescription( Tests_test_chart, suiteDescription_test_chart, 67, "test_no_write_legend" ) {}
 void runTest() { suite_test_chart.test_no_write_legend(); }
} testDescription_suite_test_chart_test_no_write_legend;

static class TestDescription_suite_test_chart_test_write_print_settings : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_chart_test_write_print_settings() : CxxTest::RealTestDescription( Tests_test_chart, suiteDescription_test_chart, 87, "test_write_print_settings" ) {}
 void runTest() { suite_test_chart.test_write_print_settings(); }
} testDescription_suite_test_chart_test_write_print_settings;

static class TestDescription_suite_test_chart_test_write_chart : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_chart_test_write_chart() : CxxTest::RealTestDescription( Tests_test_chart, suiteDescription_test_chart, 93, "test_write_chart" ) {}
 void runTest() { suite_test_chart.test_write_chart(); }
} testDescription_suite_test_chart_test_write_chart;

static class TestDescription_suite_test_chart_test_write_xaxis_scatter : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_chart_test_write_xaxis_scatter() : CxxTest::RealTestDescription( Tests_test_chart, suiteDescription_test_chart, 118, "test_write_xaxis_scatter" ) {}
 void runTest() { suite_test_chart.test_write_xaxis_scatter(); }
} testDescription_suite_test_chart_test_write_xaxis_scatter;

static class TestDescription_suite_test_chart_test_write_yaxis_scatter : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_chart_test_write_yaxis_scatter() : CxxTest::RealTestDescription( Tests_test_chart, suiteDescription_test_chart, 125, "test_write_yaxis_scatter" ) {}
 void runTest() { suite_test_chart.test_write_yaxis_scatter(); }
} testDescription_suite_test_chart_test_write_yaxis_scatter;

static class TestDescription_suite_test_chart_test_write_series_scatter : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_chart_test_write_series_scatter() : CxxTest::RealTestDescription( Tests_test_chart, suiteDescription_test_chart, 132, "test_write_series_scatter" ) {}
 void runTest() { suite_test_chart.test_write_series_scatter(); }
} testDescription_suite_test_chart_test_write_series_scatter;

static class TestDescription_suite_test_chart_test_write_legend_scatter : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_chart_test_write_legend_scatter() : CxxTest::RealTestDescription( Tests_test_chart, suiteDescription_test_chart, 138, "test_write_legend_scatter" ) {}
 void runTest() { suite_test_chart.test_write_legend_scatter(); }
} testDescription_suite_test_chart_test_write_legend_scatter;

static class TestDescription_suite_test_chart_test_write_print_settings_scatter : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_chart_test_write_print_settings_scatter() : CxxTest::RealTestDescription( Tests_test_chart, suiteDescription_test_chart, 144, "test_write_print_settings_scatter" ) {}
 void runTest() { suite_test_chart.test_write_print_settings_scatter(); }
} testDescription_suite_test_chart_test_write_print_settings_scatter;

static class TestDescription_suite_test_chart_test_write_chart_scatter : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_chart_test_write_chart_scatter() : CxxTest::RealTestDescription( Tests_test_chart, suiteDescription_test_chart, 150, "test_write_chart_scatter" ) {}
 void runTest() { suite_test_chart.test_write_chart_scatter(); }
} testDescription_suite_test_chart_test_write_chart_scatter;

#include "/Users/thomas/Development/xlnt/tests/test_dump.hpp"

static test_dump suite_test_dump;

static CxxTest::List Tests_test_dump = { 0, 0 };
CxxTest::StaticSuiteDescription suiteDescription_test_dump( "../../tests/test_dump.hpp", 9, "test_dump", suite_test_dump, Tests_test_dump );

static class TestDescription_suite_test_dump_test_dump_sheet_title : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_dump_test_dump_sheet_title() : CxxTest::RealTestDescription( Tests_test_dump, suiteDescription_test_dump, 12, "test_dump_sheet_title" ) {}
 void runTest() { suite_test_dump.test_dump_sheet_title(); }
} testDescription_suite_test_dump_test_dump_sheet_title;

static class TestDescription_suite_test_dump_test_dump_sheet : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_dump_test_dump_sheet() : CxxTest::RealTestDescription( Tests_test_dump, suiteDescription_test_dump, 24, "test_dump_sheet" ) {}
 void runTest() { suite_test_dump.test_dump_sheet(); }
} testDescription_suite_test_dump_test_dump_sheet;

static class TestDescription_suite_test_dump_test_table_builder : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_dump_test_table_builder() : CxxTest::RealTestDescription( Tests_test_dump, suiteDescription_test_dump, 116, "test_table_builder" ) {}
 void runTest() { suite_test_dump.test_table_builder(); }
} testDescription_suite_test_dump_test_table_builder;

static class TestDescription_suite_test_dump_test_dump_twice : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_dump_test_dump_twice() : CxxTest::RealTestDescription( Tests_test_dump, suiteDescription_test_dump, 137, "test_dump_twice" ) {}
 void runTest() { suite_test_dump.test_dump_twice(); }
} testDescription_suite_test_dump_test_dump_twice;

static class TestDescription_suite_test_dump_test_append_after_save : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_dump_test_append_after_save() : CxxTest::RealTestDescription( Tests_test_dump, suiteDescription_test_dump, 153, "test_append_after_save" ) {}
 void runTest() { suite_test_dump.test_append_after_save(); }
} testDescription_suite_test_dump_test_append_after_save;

#include "/Users/thomas/Development/xlnt/tests/test_named_range.hpp"

static test_named_range suite_test_named_range;

static CxxTest::List Tests_test_named_range = { 0, 0 };
CxxTest::StaticSuiteDescription suiteDescription_test_named_range( "../../tests/test_named_range.hpp", 8, "test_named_range", suite_test_named_range, Tests_test_named_range );

static class TestDescription_suite_test_named_range_test_split : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_named_range_test_split() : CxxTest::RealTestDescription( Tests_test_named_range, suiteDescription_test_named_range, 11, "test_split" ) {}
 void runTest() { suite_test_named_range.test_split(); }
} testDescription_suite_test_named_range_test_split;

static class TestDescription_suite_test_named_range_test_split_no_quotes : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_named_range_test_split_no_quotes() : CxxTest::RealTestDescription( Tests_test_named_range, suiteDescription_test_named_range, 16, "test_split_no_quotes" ) {}
 void runTest() { suite_test_named_range.test_split_no_quotes(); }
} testDescription_suite_test_named_range_test_split_no_quotes;

static class TestDescription_suite_test_named_range_test_bad_range_name : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_named_range_test_bad_range_name() : CxxTest::RealTestDescription( Tests_test_named_range, suiteDescription_test_named_range, 21, "test_bad_range_name" ) {}
 void runTest() { suite_test_named_range.test_bad_range_name(); }
} testDescription_suite_test_named_range_test_bad_range_name;

static class TestDescription_suite_test_named_range_test_range_name_worksheet_special_chars : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_named_range_test_range_name_worksheet_special_chars() : CxxTest::RealTestDescription( Tests_test_named_range, suiteDescription_test_named_range, 26, "test_range_name_worksheet_special_chars" ) {}
 void runTest() { suite_test_named_range.test_range_name_worksheet_special_chars(); }
} testDescription_suite_test_named_range_test_range_name_worksheet_special_chars;

static class TestDescription_suite_test_named_range_test_read_named_ranges : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_named_range_test_read_named_ranges() : CxxTest::RealTestDescription( Tests_test_named_range, suiteDescription_test_named_range, 65, "test_read_named_ranges" ) {}
 void runTest() { suite_test_named_range.test_read_named_ranges(); }
} testDescription_suite_test_named_range_test_read_named_ranges;

static class TestDescription_suite_test_named_range_test_oddly_shaped_named_ranges : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_named_range_test_oddly_shaped_named_ranges() : CxxTest::RealTestDescription( Tests_test_named_range, suiteDescription_test_named_range, 94, "test_oddly_shaped_named_ranges" ) {}
 void runTest() { suite_test_named_range.test_oddly_shaped_named_ranges(); }
} testDescription_suite_test_named_range_test_oddly_shaped_named_ranges;

static class TestDescription_suite_test_named_range_test_merged_cells_named_range : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_named_range_test_merged_cells_named_range() : CxxTest::RealTestDescription( Tests_test_named_range, suiteDescription_test_named_range, 117, "test_merged_cells_named_range" ) {}
 void runTest() { suite_test_named_range.test_merged_cells_named_range(); }
} testDescription_suite_test_named_range_test_merged_cells_named_range;

static class TestDescription_suite_test_named_range_test_has_ranges : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_named_range_test_has_ranges() : CxxTest::RealTestDescription( Tests_test_named_range, suiteDescription_test_named_range, 143, "test_has_ranges" ) {}
 void runTest() { suite_test_named_range.test_has_ranges(); }
} testDescription_suite_test_named_range_test_has_ranges;

static class TestDescription_suite_test_named_range_test_workbook_has_normal_range : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_named_range_test_workbook_has_normal_range() : CxxTest::RealTestDescription( Tests_test_named_range, suiteDescription_test_named_range, 149, "test_workbook_has_normal_range" ) {}
 void runTest() { suite_test_named_range.test_workbook_has_normal_range(); }
} testDescription_suite_test_named_range_test_workbook_has_normal_range;

static class TestDescription_suite_test_named_range_test_workbook_has_value_range : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_named_range_test_workbook_has_value_range() : CxxTest::RealTestDescription( Tests_test_named_range, suiteDescription_test_named_range, 155, "test_workbook_has_value_range" ) {}
 void runTest() { suite_test_named_range.test_workbook_has_value_range(); }
} testDescription_suite_test_named_range_test_workbook_has_value_range;

static class TestDescription_suite_test_named_range_test_worksheet_range : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_named_range_test_worksheet_range() : CxxTest::RealTestDescription( Tests_test_named_range, suiteDescription_test_named_range, 162, "test_worksheet_range" ) {}
 void runTest() { suite_test_named_range.test_worksheet_range(); }
} testDescription_suite_test_named_range_test_worksheet_range;

static class TestDescription_suite_test_named_range_test_worksheet_range_error_on_value_range : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_named_range_test_worksheet_range_error_on_value_range() : CxxTest::RealTestDescription( Tests_test_named_range, suiteDescription_test_named_range, 167, "test_worksheet_range_error_on_value_range" ) {}
 void runTest() { suite_test_named_range.test_worksheet_range_error_on_value_range(); }
} testDescription_suite_test_named_range_test_worksheet_range_error_on_value_range;

static class TestDescription_suite_test_named_range_test_handles_scope : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_named_range_test_handles_scope() : CxxTest::RealTestDescription( Tests_test_named_range, suiteDescription_test_named_range, 194, "test_handles_scope" ) {}
 void runTest() { suite_test_named_range.test_handles_scope(); }
} testDescription_suite_test_named_range_test_handles_scope;

static class TestDescription_suite_test_named_range_test_can_be_saved : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_named_range_test_can_be_saved() : CxxTest::RealTestDescription( Tests_test_named_range, suiteDescription_test_named_range, 201, "test_can_be_saved" ) {}
 void runTest() { suite_test_named_range.test_can_be_saved(); }
} testDescription_suite_test_named_range_test_can_be_saved;

#include "/Users/thomas/Development/xlnt/tests/test_number_format.hpp"

static test_number_format suite_test_number_format;

static CxxTest::List Tests_test_number_format = { 0, 0 };
CxxTest::StaticSuiteDescription suiteDescription_test_number_format( "../../tests/test_number_format.hpp", 8, "test_number_format", suite_test_number_format, Tests_test_number_format );

static class TestDescription_suite_test_number_format_test_convert_date_to_julian : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_number_format_test_convert_date_to_julian() : CxxTest::RealTestDescription( Tests_test_number_format, suiteDescription_test_number_format, 23, "test_convert_date_to_julian" ) {}
 void runTest() { suite_test_number_format.test_convert_date_to_julian(); }
} testDescription_suite_test_number_format_test_convert_date_to_julian;

static class TestDescription_suite_test_number_format_test_convert_date_from_julian : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_number_format_test_convert_date_from_julian() : CxxTest::RealTestDescription( Tests_test_number_format, suiteDescription_test_number_format, 28, "test_convert_date_from_julian" ) {}
 void runTest() { suite_test_number_format.test_convert_date_from_julian(); }
} testDescription_suite_test_number_format_test_convert_date_from_julian;

static class TestDescription_suite_test_number_format_test_convert_datetime_to_julian : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_number_format_test_convert_datetime_to_julian() : CxxTest::RealTestDescription( Tests_test_number_format, suiteDescription_test_number_format, 47, "test_convert_datetime_to_julian" ) {}
 void runTest() { suite_test_number_format.test_convert_datetime_to_julian(); }
} testDescription_suite_test_number_format_test_convert_datetime_to_julian;

static class TestDescription_suite_test_number_format_test_insert_float : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_number_format_test_insert_float() : CxxTest::RealTestDescription( Tests_test_number_format, suiteDescription_test_number_format, 53, "test_insert_float" ) {}
 void runTest() { suite_test_number_format.test_insert_float(); }
} testDescription_suite_test_number_format_test_insert_float;

static class TestDescription_suite_test_number_format_test_insert_percentage : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_number_format_test_insert_percentage() : CxxTest::RealTestDescription( Tests_test_number_format, suiteDescription_test_number_format, 59, "test_insert_percentage" ) {}
 void runTest() { suite_test_number_format.test_insert_percentage(); }
} testDescription_suite_test_number_format_test_insert_percentage;

static class TestDescription_suite_test_number_format_test_insert_datetime : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_number_format_test_insert_datetime() : CxxTest::RealTestDescription( Tests_test_number_format, suiteDescription_test_number_format, 66, "test_insert_datetime" ) {}
 void runTest() { suite_test_number_format.test_insert_datetime(); }
} testDescription_suite_test_number_format_test_insert_datetime;

static class TestDescription_suite_test_number_format_test_insert_date : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_number_format_test_insert_date() : CxxTest::RealTestDescription( Tests_test_number_format, suiteDescription_test_number_format, 72, "test_insert_date" ) {}
 void runTest() { suite_test_number_format.test_insert_date(); }
} testDescription_suite_test_number_format_test_insert_date;

static class TestDescription_suite_test_number_format_test_internal_date : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_number_format_test_internal_date() : CxxTest::RealTestDescription( Tests_test_number_format, suiteDescription_test_number_format, 78, "test_internal_date" ) {}
 void runTest() { suite_test_number_format.test_internal_date(); }
} testDescription_suite_test_number_format_test_internal_date;

static class TestDescription_suite_test_number_format_test_datetime_interpretation : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_number_format_test_datetime_interpretation() : CxxTest::RealTestDescription( Tests_test_number_format, suiteDescription_test_number_format, 85, "test_datetime_interpretation" ) {}
 void runTest() { suite_test_number_format.test_datetime_interpretation(); }
} testDescription_suite_test_number_format_test_datetime_interpretation;

static class TestDescription_suite_test_number_format_test_date_interpretation : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_number_format_test_date_interpretation() : CxxTest::RealTestDescription( Tests_test_number_format, suiteDescription_test_number_format, 92, "test_date_interpretation" ) {}
 void runTest() { suite_test_number_format.test_date_interpretation(); }
} testDescription_suite_test_number_format_test_date_interpretation;

static class TestDescription_suite_test_number_format_test_number_format_style : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_number_format_test_number_format_style() : CxxTest::RealTestDescription( Tests_test_number_format, suiteDescription_test_number_format, 99, "test_number_format_style" ) {}
 void runTest() { suite_test_number_format.test_number_format_style(); }
} testDescription_suite_test_number_format_test_number_format_style;

static class TestDescription_suite_test_number_format_test_date_format_on_non_date : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_number_format_test_date_format_on_non_date() : CxxTest::RealTestDescription( Tests_test_number_format, suiteDescription_test_number_format, 106, "test_date_format_on_non_date" ) {}
 void runTest() { suite_test_number_format.test_date_format_on_non_date(); }
} testDescription_suite_test_number_format_test_date_format_on_non_date;

static class TestDescription_suite_test_number_format_test_1900_leap_year : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_number_format_test_1900_leap_year() : CxxTest::RealTestDescription( Tests_test_number_format, suiteDescription_test_number_format, 128, "test_1900_leap_year" ) {}
 void runTest() { suite_test_number_format.test_1900_leap_year(); }
} testDescription_suite_test_number_format_test_1900_leap_year;

static class TestDescription_suite_test_number_format_test_bad_date : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_number_format_test_bad_date() : CxxTest::RealTestDescription( Tests_test_number_format, suiteDescription_test_number_format, 134, "test_bad_date" ) {}
 void runTest() { suite_test_number_format.test_bad_date(); }
} testDescription_suite_test_number_format_test_bad_date;

static class TestDescription_suite_test_number_format_test_bad_julian_date : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_number_format_test_bad_julian_date() : CxxTest::RealTestDescription( Tests_test_number_format, suiteDescription_test_number_format, 148, "test_bad_julian_date" ) {}
 void runTest() { suite_test_number_format.test_bad_julian_date(); }
} testDescription_suite_test_number_format_test_bad_julian_date;

static class TestDescription_suite_test_number_format_test_mac_date : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_number_format_test_mac_date() : CxxTest::RealTestDescription( Tests_test_number_format, suiteDescription_test_number_format, 153, "test_mac_date" ) {}
 void runTest() { suite_test_number_format.test_mac_date(); }
} testDescription_suite_test_number_format_test_mac_date;

#include "/Users/thomas/Development/xlnt/tests/test_props.hpp"

static test_props suite_test_props;

static CxxTest::List Tests_test_props = { 0, 0 };
CxxTest::StaticSuiteDescription suiteDescription_test_props( "../../tests/test_props.hpp", 8, "test_props", suite_test_props, Tests_test_props );

static class TestDescription_suite_test_props_test_read_properties_core : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_props_test_read_properties_core() : CxxTest::RealTestDescription( Tests_test_props, suiteDescription_test_props, 25, "test_read_properties_core" ) {}
 void runTest() { suite_test_props.test_read_properties_core(); }
} testDescription_suite_test_props_test_read_properties_core;

static class TestDescription_suite_test_props_test_read_sheets_titles : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_props_test_read_sheets_titles() : CxxTest::RealTestDescription( Tests_test_props, suiteDescription_test_props, 36, "test_read_sheets_titles" ) {}
 void runTest() { suite_test_props.test_read_sheets_titles(); }
} testDescription_suite_test_props_test_read_sheets_titles;

static class TestDescription_suite_test_props_test_read_properties_core2 : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_props_test_read_properties_core2() : CxxTest::RealTestDescription( Tests_test_props, suiteDescription_test_props, 57, "test_read_properties_core2" ) {}
 void runTest() { suite_test_props.test_read_properties_core2(); }
} testDescription_suite_test_props_test_read_properties_core2;

static class TestDescription_suite_test_props_test_read_sheets_titles2 : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_props_test_read_sheets_titles2() : CxxTest::RealTestDescription( Tests_test_props, suiteDescription_test_props, 64, "test_read_sheets_titles2" ) {}
 void runTest() { suite_test_props.test_read_sheets_titles2(); }
} testDescription_suite_test_props_test_read_sheets_titles2;

static class TestDescription_suite_test_props_test_write_properties_core : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_props_test_write_properties_core() : CxxTest::RealTestDescription( Tests_test_props, suiteDescription_test_props, 72, "test_write_properties_core" ) {}
 void runTest() { suite_test_props.test_write_properties_core(); }
} testDescription_suite_test_props_test_write_properties_core;

static class TestDescription_suite_test_props_test_write_properties_app : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_props_test_write_properties_app() : CxxTest::RealTestDescription( Tests_test_props, suiteDescription_test_props, 84, "test_write_properties_app" ) {}
 void runTest() { suite_test_props.test_write_properties_app(); }
} testDescription_suite_test_props_test_write_properties_app;

#include "/Users/thomas/Development/xlnt/tests/test_read.hpp"

static test_read suite_test_read;

static CxxTest::List Tests_test_read = { 0, 0 };
CxxTest::StaticSuiteDescription suiteDescription_test_read( "../../tests/test_read.hpp", 10, "test_read", suite_test_read, Tests_test_read );

static class TestDescription_suite_test_read_test_read_standalone_worksheet : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_read_test_read_standalone_worksheet() : CxxTest::RealTestDescription( Tests_test_read, suiteDescription_test_read, 28, "test_read_standalone_worksheet" ) {}
 void runTest() { suite_test_read.test_read_standalone_worksheet(); }
} testDescription_suite_test_read_test_read_standalone_worksheet;

static class TestDescription_suite_test_read_test_read_standard_workbook : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_read_test_read_standard_workbook() : CxxTest::RealTestDescription( Tests_test_read, suiteDescription_test_read, 46, "test_read_standard_workbook" ) {}
 void runTest() { suite_test_read.test_read_standard_workbook(); }
} testDescription_suite_test_read_test_read_standard_workbook;

static class TestDescription_suite_test_read_test_read_standard_workbook_from_fileobj : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_read_test_read_standard_workbook_from_fileobj() : CxxTest::RealTestDescription( Tests_test_read, suiteDescription_test_read, 53, "test_read_standard_workbook_from_fileobj" ) {}
 void runTest() { suite_test_read.test_read_standard_workbook_from_fileobj(); }
} testDescription_suite_test_read_test_read_standard_workbook_from_fileobj;

static class TestDescription_suite_test_read_test_read_worksheet : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_read_test_read_worksheet() : CxxTest::RealTestDescription( Tests_test_read, suiteDescription_test_read, 63, "test_read_worksheet" ) {}
 void runTest() { suite_test_read.test_read_worksheet(); }
} testDescription_suite_test_read_test_read_worksheet;

static class TestDescription_suite_test_read_test_read_nostring_workbook : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_read_test_read_nostring_workbook() : CxxTest::RealTestDescription( Tests_test_read, suiteDescription_test_read, 74, "test_read_nostring_workbook" ) {}
 void runTest() { suite_test_read.test_read_nostring_workbook(); }
} testDescription_suite_test_read_test_read_nostring_workbook;

static class TestDescription_suite_test_read_test_read_empty_file : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_read_test_read_empty_file() : CxxTest::RealTestDescription( Tests_test_read, suiteDescription_test_read, 81, "test_read_empty_file" ) {}
 void runTest() { suite_test_read.test_read_empty_file(); }
} testDescription_suite_test_read_test_read_empty_file;

static class TestDescription_suite_test_read_test_read_empty_archive : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_read_test_read_empty_archive() : CxxTest::RealTestDescription( Tests_test_read, suiteDescription_test_read, 89, "test_read_empty_archive" ) {}
 void runTest() { suite_test_read.test_read_empty_archive(); }
} testDescription_suite_test_read_test_read_empty_archive;

static class TestDescription_suite_test_read_test_read_dimension : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_read_test_read_dimension() : CxxTest::RealTestDescription( Tests_test_read, suiteDescription_test_read, 96, "test_read_dimension" ) {}
 void runTest() { suite_test_read.test_read_dimension(); }
} testDescription_suite_test_read_test_read_dimension;

static class TestDescription_suite_test_read_test_calculate_dimension_iter : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_read_test_calculate_dimension_iter() : CxxTest::RealTestDescription( Tests_test_read, suiteDescription_test_read, 106, "test_calculate_dimension_iter" ) {}
 void runTest() { suite_test_read.test_calculate_dimension_iter(); }
} testDescription_suite_test_read_test_calculate_dimension_iter;

static class TestDescription_suite_test_read_test_get_highest_row_iter : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_read_test_get_highest_row_iter() : CxxTest::RealTestDescription( Tests_test_read, suiteDescription_test_read, 116, "test_get_highest_row_iter" ) {}
 void runTest() { suite_test_read.test_get_highest_row_iter(); }
} testDescription_suite_test_read_test_get_highest_row_iter;

static class TestDescription_suite_test_read_test_read_workbook_with_no_properties : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_read_test_read_workbook_with_no_properties() : CxxTest::RealTestDescription( Tests_test_read, suiteDescription_test_read, 126, "test_read_workbook_with_no_properties" ) {}
 void runTest() { suite_test_read.test_read_workbook_with_no_properties(); }
} testDescription_suite_test_read_test_read_workbook_with_no_properties;

static class TestDescription_suite_test_read_test_read_general_style : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_read_test_read_general_style() : CxxTest::RealTestDescription( Tests_test_read, suiteDescription_test_read, 133, "test_read_general_style" ) {}
 void runTest() { suite_test_read.test_read_general_style(); }
} testDescription_suite_test_read_test_read_general_style;

static class TestDescription_suite_test_read_test_read_date_style : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_read_test_read_date_style() : CxxTest::RealTestDescription( Tests_test_read, suiteDescription_test_read, 138, "test_read_date_style" ) {}
 void runTest() { suite_test_read.test_read_date_style(); }
} testDescription_suite_test_read_test_read_date_style;

static class TestDescription_suite_test_read_test_read_number_style : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_read_test_read_number_style() : CxxTest::RealTestDescription( Tests_test_read, suiteDescription_test_read, 143, "test_read_number_style" ) {}
 void runTest() { suite_test_read.test_read_number_style(); }
} testDescription_suite_test_read_test_read_number_style;

static class TestDescription_suite_test_read_test_read_time_style : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_read_test_read_time_style() : CxxTest::RealTestDescription( Tests_test_read, suiteDescription_test_read, 148, "test_read_time_style" ) {}
 void runTest() { suite_test_read.test_read_time_style(); }
} testDescription_suite_test_read_test_read_time_style;

static class TestDescription_suite_test_read_test_read_percentage_style : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_read_test_read_percentage_style() : CxxTest::RealTestDescription( Tests_test_read, suiteDescription_test_read, 153, "test_read_percentage_style" ) {}
 void runTest() { suite_test_read.test_read_percentage_style(); }
} testDescription_suite_test_read_test_read_percentage_style;

static class TestDescription_suite_test_read_test_read_win_base_date : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_read_test_read_win_base_date() : CxxTest::RealTestDescription( Tests_test_read, suiteDescription_test_read, 158, "test_read_win_base_date" ) {}
 void runTest() { suite_test_read.test_read_win_base_date(); }
} testDescription_suite_test_read_test_read_win_base_date;

static class TestDescription_suite_test_read_test_read_mac_base_date : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_read_test_read_mac_base_date() : CxxTest::RealTestDescription( Tests_test_read, suiteDescription_test_read, 163, "test_read_mac_base_date" ) {}
 void runTest() { suite_test_read.test_read_mac_base_date(); }
} testDescription_suite_test_read_test_read_mac_base_date;

static class TestDescription_suite_test_read_test_read_date_style_mac : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_read_test_read_date_style_mac() : CxxTest::RealTestDescription( Tests_test_read, suiteDescription_test_read, 168, "test_read_date_style_mac" ) {}
 void runTest() { suite_test_read.test_read_date_style_mac(); }
} testDescription_suite_test_read_test_read_date_style_mac;

static class TestDescription_suite_test_read_test_read_date_style_win : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_read_test_read_date_style_win() : CxxTest::RealTestDescription( Tests_test_read, suiteDescription_test_read, 173, "test_read_date_style_win" ) {}
 void runTest() { suite_test_read.test_read_date_style_win(); }
} testDescription_suite_test_read_test_read_date_style_win;

static class TestDescription_suite_test_read_test_read_date_value : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_read_test_read_date_value() : CxxTest::RealTestDescription( Tests_test_read, suiteDescription_test_read, 178, "test_read_date_value" ) {}
 void runTest() { suite_test_read.test_read_date_value(); }
} testDescription_suite_test_read_test_read_date_value;

#include "/Users/thomas/Development/xlnt/tests/test_strings.hpp"

static test_strings suite_test_strings;

static CxxTest::List Tests_test_strings = { 0, 0 };
CxxTest::StaticSuiteDescription suiteDescription_test_strings( "../../tests/test_strings.hpp", 8, "test_strings", suite_test_strings, Tests_test_strings );

static class TestDescription_suite_test_strings_test_create_string_table : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_strings_test_create_string_table() : CxxTest::RealTestDescription( Tests_test_strings, suiteDescription_test_strings, 16, "test_create_string_table" ) {}
 void runTest() { suite_test_strings.test_create_string_table(); }
} testDescription_suite_test_strings_test_create_string_table;

static class TestDescription_suite_test_strings_test_read_string_table : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_strings_test_read_string_table() : CxxTest::RealTestDescription( Tests_test_strings, suiteDescription_test_strings, 27, "test_read_string_table" ) {}
 void runTest() { suite_test_strings.test_read_string_table(); }
} testDescription_suite_test_strings_test_read_string_table;

static class TestDescription_suite_test_strings_test_empty_string : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_strings_test_empty_string() : CxxTest::RealTestDescription( Tests_test_strings, suiteDescription_test_strings, 38, "test_empty_string" ) {}
 void runTest() { suite_test_strings.test_empty_string(); }
} testDescription_suite_test_strings_test_empty_string;

static class TestDescription_suite_test_strings_test_formatted_string_table : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_strings_test_formatted_string_table() : CxxTest::RealTestDescription( Tests_test_strings, suiteDescription_test_strings, 49, "test_formatted_string_table" ) {}
 void runTest() { suite_test_strings.test_formatted_string_table(); }
} testDescription_suite_test_strings_test_formatted_string_table;

#include "/Users/thomas/Development/xlnt/tests/test_style.hpp"

static test_style suite_test_style;

static CxxTest::List Tests_test_style = { 0, 0 };
CxxTest::StaticSuiteDescription suiteDescription_test_style( "../../tests/test_style.hpp", 8, "test_style", suite_test_style, Tests_test_style );

static class TestDescription_suite_test_style_test_create_style_table : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_style_test_create_style_table() : CxxTest::RealTestDescription( Tests_test_style, suiteDescription_test_style, 11, "test_create_style_table" ) {}
 void runTest() { suite_test_style.test_create_style_table(); }
} testDescription_suite_test_style_test_create_style_table;

static class TestDescription_suite_test_style_test_write_style_table : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_style_test_write_style_table() : CxxTest::RealTestDescription( Tests_test_style, suiteDescription_test_style, 16, "test_write_style_table" ) {}
 void runTest() { suite_test_style.test_write_style_table(); }
} testDescription_suite_test_style_test_write_style_table;

static class TestDescription_suite_test_style_test_no_style : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_style_test_no_style() : CxxTest::RealTestDescription( Tests_test_style, suiteDescription_test_style, 28, "test_no_style" ) {}
 void runTest() { suite_test_style.test_no_style(); }
} testDescription_suite_test_style_test_no_style;

static class TestDescription_suite_test_style_test_nb_style : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_style_test_nb_style() : CxxTest::RealTestDescription( Tests_test_style, suiteDescription_test_style, 35, "test_nb_style" ) {}
 void runTest() { suite_test_style.test_nb_style(); }
} testDescription_suite_test_style_test_nb_style;

static class TestDescription_suite_test_style_test_style_unicity : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_style_test_style_unicity() : CxxTest::RealTestDescription( Tests_test_style, suiteDescription_test_style, 49, "test_style_unicity" ) {}
 void runTest() { suite_test_style.test_style_unicity(); }
} testDescription_suite_test_style_test_style_unicity;

static class TestDescription_suite_test_style_test_fonts : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_style_test_fonts() : CxxTest::RealTestDescription( Tests_test_style, suiteDescription_test_style, 59, "test_fonts" ) {}
 void runTest() { suite_test_style.test_fonts(); }
} testDescription_suite_test_style_test_fonts;

static class TestDescription_suite_test_style_test_fills : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_style_test_fills() : CxxTest::RealTestDescription( Tests_test_style, suiteDescription_test_style, 73, "test_fills" ) {}
 void runTest() { suite_test_style.test_fills(); }
} testDescription_suite_test_style_test_fills;

static class TestDescription_suite_test_style_test_borders : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_style_test_borders() : CxxTest::RealTestDescription( Tests_test_style, suiteDescription_test_style, 82, "test_borders" ) {}
 void runTest() { suite_test_style.test_borders(); }
} testDescription_suite_test_style_test_borders;

static class TestDescription_suite_test_style_test_write_cell_xfs_1 : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_style_test_write_cell_xfs_1() : CxxTest::RealTestDescription( Tests_test_style, suiteDescription_test_style, 91, "test_write_cell_xfs_1" ) {}
 void runTest() { suite_test_style.test_write_cell_xfs_1(); }
} testDescription_suite_test_style_test_write_cell_xfs_1;

static class TestDescription_suite_test_style_test_alignment : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_style_test_alignment() : CxxTest::RealTestDescription( Tests_test_style, suiteDescription_test_style, 105, "test_alignment" ) {}
 void runTest() { suite_test_style.test_alignment(); }
} testDescription_suite_test_style_test_alignment;

static class TestDescription_suite_test_style_test_alignment_rotation : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_style_test_alignment_rotation() : CxxTest::RealTestDescription( Tests_test_style, suiteDescription_test_style, 118, "test_alignment_rotation" ) {}
 void runTest() { suite_test_style.test_alignment_rotation(); }
} testDescription_suite_test_style_test_alignment_rotation;

static class TestDescription_suite_test_style_test_format_comparisions : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_style_test_format_comparisions() : CxxTest::RealTestDescription( Tests_test_style, suiteDescription_test_style, 133, "test_format_comparisions" ) {}
 void runTest() { suite_test_style.test_format_comparisions(); }
} testDescription_suite_test_style_test_format_comparisions;

static class TestDescription_suite_test_style_test_builtin_format : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_style_test_builtin_format() : CxxTest::RealTestDescription( Tests_test_style, suiteDescription_test_style, 148, "test_builtin_format" ) {}
 void runTest() { suite_test_style.test_builtin_format(); }
} testDescription_suite_test_style_test_builtin_format;

static class TestDescription_suite_test_style_test_read_style : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_style_test_read_style() : CxxTest::RealTestDescription( Tests_test_style, suiteDescription_test_style, 156, "test_read_style" ) {}
 void runTest() { suite_test_style.test_read_style(); }
} testDescription_suite_test_style_test_read_style;

static class TestDescription_suite_test_style_test_read_cell_style : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_style_test_read_cell_style() : CxxTest::RealTestDescription( Tests_test_style, suiteDescription_test_style, 173, "test_read_cell_style" ) {}
 void runTest() { suite_test_style.test_read_cell_style(); }
} testDescription_suite_test_style_test_read_cell_style;

#include "/Users/thomas/Development/xlnt/tests/test_theme.hpp"

static test_theme suite_test_theme;

static CxxTest::List Tests_test_theme = { 0, 0 };
CxxTest::StaticSuiteDescription suiteDescription_test_theme( "../../tests/test_theme.hpp", 10, "test_theme", suite_test_theme, Tests_test_theme );

static class TestDescription_suite_test_theme_test_write_theme : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_theme_test_write_theme() : CxxTest::RealTestDescription( Tests_test_theme, suiteDescription_test_theme, 13, "test_write_theme" ) {}
 void runTest() { suite_test_theme.test_write_theme(); }
} testDescription_suite_test_theme_test_write_theme;

#include "/Users/thomas/Development/xlnt/tests/test_workbook.hpp"

static test_workbook suite_test_workbook;

static CxxTest::List Tests_test_workbook = { 0, 0 };
CxxTest::StaticSuiteDescription suiteDescription_test_workbook( "../../tests/test_workbook.hpp", 9, "test_workbook", suite_test_workbook, Tests_test_workbook );

static class TestDescription_suite_test_workbook_test_get_active_sheet : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_workbook_test_get_active_sheet() : CxxTest::RealTestDescription( Tests_test_workbook, suiteDescription_test_workbook, 12, "test_get_active_sheet" ) {}
 void runTest() { suite_test_workbook.test_get_active_sheet(); }
} testDescription_suite_test_workbook_test_get_active_sheet;

static class TestDescription_suite_test_workbook_test_create_sheet : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_workbook_test_create_sheet() : CxxTest::RealTestDescription( Tests_test_workbook, suiteDescription_test_workbook, 19, "test_create_sheet" ) {}
 void runTest() { suite_test_workbook.test_create_sheet(); }
} testDescription_suite_test_workbook_test_create_sheet;

static class TestDescription_suite_test_workbook_test_create_sheet_with_name : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_workbook_test_create_sheet_with_name() : CxxTest::RealTestDescription( Tests_test_workbook, suiteDescription_test_workbook, 26, "test_create_sheet_with_name" ) {}
 void runTest() { suite_test_workbook.test_create_sheet_with_name(); }
} testDescription_suite_test_workbook_test_create_sheet_with_name;

static class TestDescription_suite_test_workbook_test_create_sheet_readonly : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_workbook_test_create_sheet_readonly() : CxxTest::RealTestDescription( Tests_test_workbook, suiteDescription_test_workbook, 33, "test_create_sheet_readonly" ) {}
 void runTest() { suite_test_workbook.test_create_sheet_readonly(); }
} testDescription_suite_test_workbook_test_create_sheet_readonly;

static class TestDescription_suite_test_workbook_test_remove_sheet : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_workbook_test_remove_sheet() : CxxTest::RealTestDescription( Tests_test_workbook, suiteDescription_test_workbook, 39, "test_remove_sheet" ) {}
 void runTest() { suite_test_workbook.test_remove_sheet(); }
} testDescription_suite_test_workbook_test_remove_sheet;

static class TestDescription_suite_test_workbook_test_get_sheet_by_name : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_workbook_test_get_sheet_by_name() : CxxTest::RealTestDescription( Tests_test_workbook, suiteDescription_test_workbook, 51, "test_get_sheet_by_name" ) {}
 void runTest() { suite_test_workbook.test_get_sheet_by_name(); }
} testDescription_suite_test_workbook_test_get_sheet_by_name;

static class TestDescription_suite_test_workbook_test_getitem : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_workbook_test_getitem() : CxxTest::RealTestDescription( Tests_test_workbook, suiteDescription_test_workbook, 61, "test_getitem" ) {}
 void runTest() { suite_test_workbook.test_getitem(); }
} testDescription_suite_test_workbook_test_getitem;

static class TestDescription_suite_test_workbook_test_get_index2 : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_workbook_test_get_index2() : CxxTest::RealTestDescription( Tests_test_workbook, suiteDescription_test_workbook, 73, "test_get_index2" ) {}
 void runTest() { suite_test_workbook.test_get_index2(); }
} testDescription_suite_test_workbook_test_get_index2;

static class TestDescription_suite_test_workbook_test_get_sheet_names : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_workbook_test_get_sheet_names() : CxxTest::RealTestDescription( Tests_test_workbook, suiteDescription_test_workbook, 81, "test_get_sheet_names" ) {}
 void runTest() { suite_test_workbook.test_get_sheet_names(); }
} testDescription_suite_test_workbook_test_get_sheet_names;

static class TestDescription_suite_test_workbook_test_get_active_sheet2 : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_workbook_test_get_active_sheet2() : CxxTest::RealTestDescription( Tests_test_workbook, suiteDescription_test_workbook, 99, "test_get_active_sheet2" ) {}
 void runTest() { suite_test_workbook.test_get_active_sheet2(); }
} testDescription_suite_test_workbook_test_get_active_sheet2;

static class TestDescription_suite_test_workbook_test_create_sheet2 : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_workbook_test_create_sheet2() : CxxTest::RealTestDescription( Tests_test_workbook, suiteDescription_test_workbook, 106, "test_create_sheet2" ) {}
 void runTest() { suite_test_workbook.test_create_sheet2(); }
} testDescription_suite_test_workbook_test_create_sheet2;

static class TestDescription_suite_test_workbook_test_create_sheet_with_name2 : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_workbook_test_create_sheet_with_name2() : CxxTest::RealTestDescription( Tests_test_workbook, suiteDescription_test_workbook, 113, "test_create_sheet_with_name2" ) {}
 void runTest() { suite_test_workbook.test_create_sheet_with_name2(); }
} testDescription_suite_test_workbook_test_create_sheet_with_name2;

static class TestDescription_suite_test_workbook_test_get_sheet_by_name2 : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_workbook_test_get_sheet_by_name2() : CxxTest::RealTestDescription( Tests_test_workbook, suiteDescription_test_workbook, 120, "test_get_sheet_by_name2" ) {}
 void runTest() { suite_test_workbook.test_get_sheet_by_name2(); }
} testDescription_suite_test_workbook_test_get_sheet_by_name2;

static class TestDescription_suite_test_workbook_test_get_index : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_workbook_test_get_index() : CxxTest::RealTestDescription( Tests_test_workbook, suiteDescription_test_workbook, 130, "test_get_index" ) {}
 void runTest() { suite_test_workbook.test_get_index(); }
} testDescription_suite_test_workbook_test_get_index;

static class TestDescription_suite_test_workbook_test_add_named_range : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_workbook_test_add_named_range() : CxxTest::RealTestDescription( Tests_test_workbook, suiteDescription_test_workbook, 138, "test_add_named_range" ) {}
 void runTest() { suite_test_workbook.test_add_named_range(); }
} testDescription_suite_test_workbook_test_add_named_range;

static class TestDescription_suite_test_workbook_test_get_named_range2 : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_workbook_test_get_named_range2() : CxxTest::RealTestDescription( Tests_test_workbook, suiteDescription_test_workbook, 147, "test_get_named_range2" ) {}
 void runTest() { suite_test_workbook.test_get_named_range2(); }
} testDescription_suite_test_workbook_test_get_named_range2;

static class TestDescription_suite_test_workbook_test_remove_named_range : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_workbook_test_remove_named_range() : CxxTest::RealTestDescription( Tests_test_workbook, suiteDescription_test_workbook, 157, "test_remove_named_range" ) {}
 void runTest() { suite_test_workbook.test_remove_named_range(); }
} testDescription_suite_test_workbook_test_remove_named_range;

static class TestDescription_suite_test_workbook_test_add_local_named_range : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_workbook_test_add_local_named_range() : CxxTest::RealTestDescription( Tests_test_workbook, suiteDescription_test_workbook, 167, "test_add_local_named_range" ) {}
 void runTest() { suite_test_workbook.test_add_local_named_range(); }
} testDescription_suite_test_workbook_test_add_local_named_range;

static class TestDescription_suite_test_workbook_test_write_regular_date : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_workbook_test_write_regular_date() : CxxTest::RealTestDescription( Tests_test_workbook, suiteDescription_test_workbook, 177, "test_write_regular_date" ) {}
 void runTest() { suite_test_workbook.test_write_regular_date(); }
} testDescription_suite_test_workbook_test_write_regular_date;

static class TestDescription_suite_test_workbook_test_write_regular_float : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_workbook_test_write_regular_float() : CxxTest::RealTestDescription( Tests_test_workbook, suiteDescription_test_workbook, 195, "test_write_regular_float" ) {}
 void runTest() { suite_test_workbook.test_write_regular_float(); }
} testDescription_suite_test_workbook_test_write_regular_float;

#include "/Users/thomas/Development/xlnt/tests/test_worksheet.hpp"

static test_worksheet suite_test_worksheet;

static CxxTest::List Tests_test_worksheet = { 0, 0 };
CxxTest::StaticSuiteDescription suiteDescription_test_worksheet( "../../tests/test_worksheet.hpp", 9, "test_worksheet", suite_test_worksheet, Tests_test_worksheet );

static class TestDescription_suite_test_worksheet_test_new_worksheet : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_worksheet_test_new_worksheet() : CxxTest::RealTestDescription( Tests_test_worksheet, suiteDescription_test_worksheet, 12, "test_new_worksheet" ) {}
 void runTest() { suite_test_worksheet.test_new_worksheet(); }
} testDescription_suite_test_worksheet_test_new_worksheet;

static class TestDescription_suite_test_worksheet_test_new_sheet_name : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_worksheet_test_new_sheet_name() : CxxTest::RealTestDescription( Tests_test_worksheet, suiteDescription_test_worksheet, 18, "test_new_sheet_name" ) {}
 void runTest() { suite_test_worksheet.test_new_sheet_name(); }
} testDescription_suite_test_worksheet_test_new_sheet_name;

static class TestDescription_suite_test_worksheet_test_get_cell : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_worksheet_test_get_cell() : CxxTest::RealTestDescription( Tests_test_worksheet, suiteDescription_test_worksheet, 26, "test_get_cell" ) {}
 void runTest() { suite_test_worksheet.test_get_cell(); }
} testDescription_suite_test_worksheet_test_get_cell;

static class TestDescription_suite_test_worksheet_test_set_bad_title : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_worksheet_test_set_bad_title() : CxxTest::RealTestDescription( Tests_test_worksheet, suiteDescription_test_worksheet, 33, "test_set_bad_title" ) {}
 void runTest() { suite_test_worksheet.test_set_bad_title(); }
} testDescription_suite_test_worksheet_test_set_bad_title;

static class TestDescription_suite_test_worksheet_test_set_bad_title_character : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_worksheet_test_set_bad_title_character() : CxxTest::RealTestDescription( Tests_test_worksheet, suiteDescription_test_worksheet, 39, "test_set_bad_title_character" ) {}
 void runTest() { suite_test_worksheet.test_set_bad_title_character(); }
} testDescription_suite_test_worksheet_test_set_bad_title_character;

static class TestDescription_suite_test_worksheet_test_worksheet_dimension : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_worksheet_test_worksheet_dimension() : CxxTest::RealTestDescription( Tests_test_worksheet, suiteDescription_test_worksheet, 50, "test_worksheet_dimension" ) {}
 void runTest() { suite_test_worksheet.test_worksheet_dimension(); }
} testDescription_suite_test_worksheet_test_worksheet_dimension;

static class TestDescription_suite_test_worksheet_test_worksheet_range : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_worksheet_test_worksheet_range() : CxxTest::RealTestDescription( Tests_test_worksheet, suiteDescription_test_worksheet, 59, "test_worksheet_range" ) {}
 void runTest() { suite_test_worksheet.test_worksheet_range(); }
} testDescription_suite_test_worksheet_test_worksheet_range;

static class TestDescription_suite_test_worksheet_test_worksheet_named_range : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_worksheet_test_worksheet_named_range() : CxxTest::RealTestDescription( Tests_test_worksheet, suiteDescription_test_worksheet, 67, "test_worksheet_named_range" ) {}
 void runTest() { suite_test_worksheet.test_worksheet_named_range(); }
} testDescription_suite_test_worksheet_test_worksheet_named_range;

static class TestDescription_suite_test_worksheet_test_bad_named_range : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_worksheet_test_bad_named_range() : CxxTest::RealTestDescription( Tests_test_worksheet, suiteDescription_test_worksheet, 77, "test_bad_named_range" ) {}
 void runTest() { suite_test_worksheet.test_bad_named_range(); }
} testDescription_suite_test_worksheet_test_bad_named_range;

static class TestDescription_suite_test_worksheet_test_named_range_wrong_sheet : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_worksheet_test_named_range_wrong_sheet() : CxxTest::RealTestDescription( Tests_test_worksheet, suiteDescription_test_worksheet, 83, "test_named_range_wrong_sheet" ) {}
 void runTest() { suite_test_worksheet.test_named_range_wrong_sheet(); }
} testDescription_suite_test_worksheet_test_named_range_wrong_sheet;

static class TestDescription_suite_test_worksheet_test_cell_offset : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_worksheet_test_cell_offset() : CxxTest::RealTestDescription( Tests_test_worksheet, suiteDescription_test_worksheet, 91, "test_cell_offset" ) {}
 void runTest() { suite_test_worksheet.test_cell_offset(); }
} testDescription_suite_test_worksheet_test_cell_offset;

static class TestDescription_suite_test_worksheet_test_range_offset : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_worksheet_test_range_offset() : CxxTest::RealTestDescription( Tests_test_worksheet, suiteDescription_test_worksheet, 97, "test_range_offset" ) {}
 void runTest() { suite_test_worksheet.test_range_offset(); }
} testDescription_suite_test_worksheet_test_range_offset;

static class TestDescription_suite_test_worksheet_test_cell_alternate_coordinates : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_worksheet_test_cell_alternate_coordinates() : CxxTest::RealTestDescription( Tests_test_worksheet, suiteDescription_test_worksheet, 106, "test_cell_alternate_coordinates" ) {}
 void runTest() { suite_test_worksheet.test_cell_alternate_coordinates(); }
} testDescription_suite_test_worksheet_test_cell_alternate_coordinates;

static class TestDescription_suite_test_worksheet_test_cell_range_name : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_worksheet_test_cell_range_name() : CxxTest::RealTestDescription( Tests_test_worksheet, suiteDescription_test_worksheet, 113, "test_cell_range_name" ) {}
 void runTest() { suite_test_worksheet.test_cell_range_name(); }
} testDescription_suite_test_worksheet_test_cell_range_name;

static class TestDescription_suite_test_worksheet_test_garbage_collect : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_worksheet_test_garbage_collect() : CxxTest::RealTestDescription( Tests_test_worksheet, suiteDescription_test_worksheet, 125, "test_garbage_collect" ) {}
 void runTest() { suite_test_worksheet.test_garbage_collect(); }
} testDescription_suite_test_worksheet_test_garbage_collect;

static class TestDescription_suite_test_worksheet_test_hyperlink_relationships : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_worksheet_test_hyperlink_relationships() : CxxTest::RealTestDescription( Tests_test_worksheet, suiteDescription_test_worksheet, 147, "test_hyperlink_relationships" ) {}
 void runTest() { suite_test_worksheet.test_hyperlink_relationships(); }
} testDescription_suite_test_worksheet_test_hyperlink_relationships;

static class TestDescription_suite_test_worksheet_test_bad_relationship_type : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_worksheet_test_bad_relationship_type() : CxxTest::RealTestDescription( Tests_test_worksheet, suiteDescription_test_worksheet, 167, "test_bad_relationship_type" ) {}
 void runTest() { suite_test_worksheet.test_bad_relationship_type(); }
} testDescription_suite_test_worksheet_test_bad_relationship_type;

static class TestDescription_suite_test_worksheet_test_append_list : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_worksheet_test_append_list() : CxxTest::RealTestDescription( Tests_test_worksheet, suiteDescription_test_worksheet, 172, "test_append_list" ) {}
 void runTest() { suite_test_worksheet.test_append_list(); }
} testDescription_suite_test_worksheet_test_append_list;

static class TestDescription_suite_test_worksheet_test_append_dict_letter : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_worksheet_test_append_dict_letter() : CxxTest::RealTestDescription( Tests_test_worksheet, suiteDescription_test_worksheet, 182, "test_append_dict_letter" ) {}
 void runTest() { suite_test_worksheet.test_append_dict_letter(); }
} testDescription_suite_test_worksheet_test_append_dict_letter;

static class TestDescription_suite_test_worksheet_test_append_dict_index : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_worksheet_test_append_dict_index() : CxxTest::RealTestDescription( Tests_test_worksheet, suiteDescription_test_worksheet, 192, "test_append_dict_index" ) {}
 void runTest() { suite_test_worksheet.test_append_dict_index(); }
} testDescription_suite_test_worksheet_test_append_dict_index;

static class TestDescription_suite_test_worksheet_test_append_2d_list : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_worksheet_test_append_2d_list() : CxxTest::RealTestDescription( Tests_test_worksheet, suiteDescription_test_worksheet, 202, "test_append_2d_list" ) {}
 void runTest() { suite_test_worksheet.test_append_2d_list(); }
} testDescription_suite_test_worksheet_test_append_2d_list;

static class TestDescription_suite_test_worksheet_test_rows : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_worksheet_test_rows() : CxxTest::RealTestDescription( Tests_test_worksheet, suiteDescription_test_worksheet, 217, "test_rows" ) {}
 void runTest() { suite_test_worksheet.test_rows(); }
} testDescription_suite_test_worksheet_test_rows;

static class TestDescription_suite_test_worksheet_test_auto_filter : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_worksheet_test_auto_filter() : CxxTest::RealTestDescription( Tests_test_worksheet, suiteDescription_test_worksheet, 232, "test_auto_filter" ) {}
 void runTest() { suite_test_worksheet.test_auto_filter(); }
} testDescription_suite_test_worksheet_test_auto_filter;

static class TestDescription_suite_test_worksheet_test_page_margins : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_worksheet_test_page_margins() : CxxTest::RealTestDescription( Tests_test_worksheet, suiteDescription_test_worksheet, 246, "test_page_margins" ) {}
 void runTest() { suite_test_worksheet.test_page_margins(); }
} testDescription_suite_test_worksheet_test_page_margins;

static class TestDescription_suite_test_worksheet_test_merge : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_worksheet_test_merge() : CxxTest::RealTestDescription( Tests_test_worksheet, suiteDescription_test_worksheet, 292, "test_merge" ) {}
 void runTest() { suite_test_worksheet.test_merge(); }
} testDescription_suite_test_worksheet_test_merge;

static class TestDescription_suite_test_worksheet_test_freeze : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_worksheet_test_freeze() : CxxTest::RealTestDescription( Tests_test_worksheet, suiteDescription_test_worksheet, 361, "test_freeze" ) {}
 void runTest() { suite_test_worksheet.test_freeze(); }
} testDescription_suite_test_worksheet_test_freeze;

static class TestDescription_suite_test_worksheet_test_printer_settings : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_worksheet_test_printer_settings() : CxxTest::RealTestDescription( Tests_test_worksheet, suiteDescription_test_worksheet, 378, "test_printer_settings" ) {}
 void runTest() { suite_test_worksheet.test_printer_settings(); }
} testDescription_suite_test_worksheet_test_printer_settings;

#include "/Users/thomas/Development/xlnt/tests/test_write.hpp"

static test_write suite_test_write;

static CxxTest::List Tests_test_write = { 0, 0 };
CxxTest::StaticSuiteDescription suiteDescription_test_write( "../../tests/test_write.hpp", 11, "test_write", suite_test_write, Tests_test_write );

static class TestDescription_suite_test_write_test_write_empty_workbook : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_write_test_write_empty_workbook() : CxxTest::RealTestDescription( Tests_test_write, suiteDescription_test_write, 14, "test_write_empty_workbook" ) {}
 void runTest() { suite_test_write.test_write_empty_workbook(); }
} testDescription_suite_test_write_test_write_empty_workbook;

static class TestDescription_suite_test_write_test_write_virtual_workbook : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_write_test_write_virtual_workbook() : CxxTest::RealTestDescription( Tests_test_write, suiteDescription_test_write, 26, "test_write_virtual_workbook" ) {}
 void runTest() { suite_test_write.test_write_virtual_workbook(); }
} testDescription_suite_test_write_test_write_virtual_workbook;

static class TestDescription_suite_test_write_test_write_workbook_rels : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_write_test_write_workbook_rels() : CxxTest::RealTestDescription( Tests_test_write, suiteDescription_test_write, 35, "test_write_workbook_rels" ) {}
 void runTest() { suite_test_write.test_write_workbook_rels(); }
} testDescription_suite_test_write_test_write_workbook_rels;

static class TestDescription_suite_test_write_test_write_workbook : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_write_test_write_workbook() : CxxTest::RealTestDescription( Tests_test_write, suiteDescription_test_write, 43, "test_write_workbook" ) {}
 void runTest() { suite_test_write.test_write_workbook(); }
} testDescription_suite_test_write_test_write_workbook;

static class TestDescription_suite_test_write_test_write_string_table : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_write_test_write_string_table() : CxxTest::RealTestDescription( Tests_test_write, suiteDescription_test_write, 51, "test_write_string_table" ) {}
 void runTest() { suite_test_write.test_write_string_table(); }
} testDescription_suite_test_write_test_write_string_table;

static class TestDescription_suite_test_write_test_write_worksheet : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_write_test_write_worksheet() : CxxTest::RealTestDescription( Tests_test_write, suiteDescription_test_write, 58, "test_write_worksheet" ) {}
 void runTest() { suite_test_write.test_write_worksheet(); }
} testDescription_suite_test_write_test_write_worksheet;

static class TestDescription_suite_test_write_test_write_hidden_worksheet : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_write_test_write_hidden_worksheet() : CxxTest::RealTestDescription( Tests_test_write, suiteDescription_test_write, 66, "test_write_hidden_worksheet" ) {}
 void runTest() { suite_test_write.test_write_hidden_worksheet(); }
} testDescription_suite_test_write_test_write_hidden_worksheet;

static class TestDescription_suite_test_write_test_write_bool : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_write_test_write_bool() : CxxTest::RealTestDescription( Tests_test_write, suiteDescription_test_write, 75, "test_write_bool" ) {}
 void runTest() { suite_test_write.test_write_bool(); }
} testDescription_suite_test_write_test_write_bool;

static class TestDescription_suite_test_write_test_write_formula : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_write_test_write_formula() : CxxTest::RealTestDescription( Tests_test_write, suiteDescription_test_write, 84, "test_write_formula" ) {}
 void runTest() { suite_test_write.test_write_formula(); }
} testDescription_suite_test_write_test_write_formula;

static class TestDescription_suite_test_write_test_write_style : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_write_test_write_style() : CxxTest::RealTestDescription( Tests_test_write, suiteDescription_test_write, 98, "test_write_style" ) {}
 void runTest() { suite_test_write.test_write_style(); }
} testDescription_suite_test_write_test_write_style;

static class TestDescription_suite_test_write_test_write_height : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_write_test_write_height() : CxxTest::RealTestDescription( Tests_test_write, suiteDescription_test_write, 108, "test_write_height" ) {}
 void runTest() { suite_test_write.test_write_height(); }
} testDescription_suite_test_write_test_write_height;

static class TestDescription_suite_test_write_test_write_hyperlink : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_write_test_write_hyperlink() : CxxTest::RealTestDescription( Tests_test_write, suiteDescription_test_write, 117, "test_write_hyperlink" ) {}
 void runTest() { suite_test_write.test_write_hyperlink(); }
} testDescription_suite_test_write_test_write_hyperlink;

static class TestDescription_suite_test_write_test_write_hyperlink_rels : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_write_test_write_hyperlink_rels() : CxxTest::RealTestDescription( Tests_test_write, suiteDescription_test_write, 126, "test_write_hyperlink_rels" ) {}
 void runTest() { suite_test_write.test_write_hyperlink_rels(); }
} testDescription_suite_test_write_test_write_hyperlink_rels;

static class TestDescription_suite_test_write_test_hyperlink_value : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_write_test_hyperlink_value() : CxxTest::RealTestDescription( Tests_test_write, suiteDescription_test_write, 140, "test_hyperlink_value" ) {}
 void runTest() { suite_test_write.test_hyperlink_value(); }
} testDescription_suite_test_write_test_hyperlink_value;

static class TestDescription_suite_test_write_test_write_auto_filter : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_write_test_write_auto_filter() : CxxTest::RealTestDescription( Tests_test_write, suiteDescription_test_write, 149, "test_write_auto_filter" ) {}
 void runTest() { suite_test_write.test_write_auto_filter(); }
} testDescription_suite_test_write_test_write_auto_filter;

static class TestDescription_suite_test_write_test_freeze_panes_horiz : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_write_test_freeze_panes_horiz() : CxxTest::RealTestDescription( Tests_test_write, suiteDescription_test_write, 162, "test_freeze_panes_horiz" ) {}
 void runTest() { suite_test_write.test_freeze_panes_horiz(); }
} testDescription_suite_test_write_test_freeze_panes_horiz;

static class TestDescription_suite_test_write_test_freeze_panes_vert : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_write_test_freeze_panes_vert() : CxxTest::RealTestDescription( Tests_test_write, suiteDescription_test_write, 171, "test_freeze_panes_vert" ) {}
 void runTest() { suite_test_write.test_freeze_panes_vert(); }
} testDescription_suite_test_write_test_freeze_panes_vert;

static class TestDescription_suite_test_write_test_freeze_panes_both : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_write_test_freeze_panes_both() : CxxTest::RealTestDescription( Tests_test_write, suiteDescription_test_write, 180, "test_freeze_panes_both" ) {}
 void runTest() { suite_test_write.test_freeze_panes_both(); }
} testDescription_suite_test_write_test_freeze_panes_both;

static class TestDescription_suite_test_write_test_long_number : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_write_test_long_number() : CxxTest::RealTestDescription( Tests_test_write, suiteDescription_test_write, 189, "test_long_number" ) {}
 void runTest() { suite_test_write.test_long_number(); }
} testDescription_suite_test_write_test_long_number;

static class TestDescription_suite_test_write_test_short_number : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_test_write_test_short_number() : CxxTest::RealTestDescription( Tests_test_write, suiteDescription_test_write, 197, "test_short_number" ) {}
 void runTest() { suite_test_write.test_short_number(); }
} testDescription_suite_test_write_test_short_number;

#include <cxxtest/Root.cpp>
const char* CxxTest::RealWorldDescription::_worldName = "cxxtest";
