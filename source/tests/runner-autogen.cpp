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
bool suite_CellTestSuite_init = false;
#include "/Users/thomas/Development/xlnt/source/tests/CellTestSuite.h"

static CellTestSuite suite_CellTestSuite;

static CxxTest::List Tests_CellTestSuite = { 0, 0 };
CxxTest::StaticSuiteDescription suiteDescription_CellTestSuite( "../../source/tests/CellTestSuite.h", 9, "CellTestSuite", suite_CellTestSuite, Tests_CellTestSuite );

static class TestDescription_suite_CellTestSuite_test_coordinates : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_CellTestSuite_test_coordinates() : CxxTest::RealTestDescription( Tests_CellTestSuite, suiteDescription_CellTestSuite, 17, "test_coordinates" ) {}
 void runTest() { suite_CellTestSuite.test_coordinates(); }
} testDescription_suite_CellTestSuite_test_coordinates;

static class TestDescription_suite_CellTestSuite_test_invalid_coordinate : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_CellTestSuite_test_invalid_coordinate() : CxxTest::RealTestDescription( Tests_CellTestSuite, suiteDescription_CellTestSuite, 24, "test_invalid_coordinate" ) {}
 void runTest() { suite_CellTestSuite.test_invalid_coordinate(); }
} testDescription_suite_CellTestSuite_test_invalid_coordinate;

static class TestDescription_suite_CellTestSuite_test_zero_row : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_CellTestSuite_test_zero_row() : CxxTest::RealTestDescription( Tests_CellTestSuite, suiteDescription_CellTestSuite, 29, "test_zero_row" ) {}
 void runTest() { suite_CellTestSuite.test_zero_row(); }
} testDescription_suite_CellTestSuite_test_zero_row;

static class TestDescription_suite_CellTestSuite_test_absolute : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_CellTestSuite_test_absolute() : CxxTest::RealTestDescription( Tests_CellTestSuite, suiteDescription_CellTestSuite, 34, "test_absolute" ) {}
 void runTest() { suite_CellTestSuite.test_absolute(); }
} testDescription_suite_CellTestSuite_test_absolute;

static class TestDescription_suite_CellTestSuite_test_absolute_multiple : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_CellTestSuite_test_absolute_multiple() : CxxTest::RealTestDescription( Tests_CellTestSuite, suiteDescription_CellTestSuite, 39, "test_absolute_multiple" ) {}
 void runTest() { suite_CellTestSuite.test_absolute_multiple(); }
} testDescription_suite_CellTestSuite_test_absolute_multiple;

static class TestDescription_suite_CellTestSuite_test_column_index : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_CellTestSuite_test_column_index() : CxxTest::RealTestDescription( Tests_CellTestSuite, suiteDescription_CellTestSuite, 44, "test_column_index" ) {}
 void runTest() { suite_CellTestSuite.test_column_index(); }
} testDescription_suite_CellTestSuite_test_column_index;

static class TestDescription_suite_CellTestSuite_test_bad_column_index : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_CellTestSuite_test_bad_column_index() : CxxTest::RealTestDescription( Tests_CellTestSuite, suiteDescription_CellTestSuite, 65, "test_bad_column_index" ) {}
 void runTest() { suite_CellTestSuite.test_bad_column_index(); }
} testDescription_suite_CellTestSuite_test_bad_column_index;

static class TestDescription_suite_CellTestSuite_test_column_letter_boundries : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_CellTestSuite_test_column_letter_boundries() : CxxTest::RealTestDescription( Tests_CellTestSuite, suiteDescription_CellTestSuite, 73, "test_column_letter_boundries" ) {}
 void runTest() { suite_CellTestSuite.test_column_letter_boundries(); }
} testDescription_suite_CellTestSuite_test_column_letter_boundries;

static class TestDescription_suite_CellTestSuite_test_column_letter : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_CellTestSuite_test_column_letter() : CxxTest::RealTestDescription( Tests_CellTestSuite, suiteDescription_CellTestSuite, 80, "test_column_letter" ) {}
 void runTest() { suite_CellTestSuite.test_column_letter(); }
} testDescription_suite_CellTestSuite_test_column_letter;

static class TestDescription_suite_CellTestSuite_test_initial_value : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_CellTestSuite_test_initial_value() : CxxTest::RealTestDescription( Tests_CellTestSuite, suiteDescription_CellTestSuite, 90, "test_initial_value" ) {}
 void runTest() { suite_CellTestSuite.test_initial_value(); }
} testDescription_suite_CellTestSuite_test_initial_value;

static class TestDescription_suite_CellTestSuite_test_null : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_CellTestSuite_test_null() : CxxTest::RealTestDescription( Tests_CellTestSuite, suiteDescription_CellTestSuite, 99, "test_null" ) {}
 void runTest() { suite_CellTestSuite.test_null(); }
} testDescription_suite_CellTestSuite_test_null;

static class TestDescription_suite_CellTestSuite_test_numeric : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_CellTestSuite_test_numeric() : CxxTest::RealTestDescription( Tests_CellTestSuite, suiteDescription_CellTestSuite, 108, "test_numeric" ) {}
 void runTest() { suite_CellTestSuite.test_numeric(); }
} testDescription_suite_CellTestSuite_test_numeric;

static class TestDescription_suite_CellTestSuite_test_string : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_CellTestSuite_test_string() : CxxTest::RealTestDescription( Tests_CellTestSuite, suiteDescription_CellTestSuite, 140, "test_string" ) {}
 void runTest() { suite_CellTestSuite.test_string(); }
} testDescription_suite_CellTestSuite_test_string;

static class TestDescription_suite_CellTestSuite_test_single_dot : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_CellTestSuite_test_single_dot() : CxxTest::RealTestDescription( Tests_CellTestSuite, suiteDescription_CellTestSuite, 150, "test_single_dot" ) {}
 void runTest() { suite_CellTestSuite.test_single_dot(); }
} testDescription_suite_CellTestSuite_test_single_dot;

static class TestDescription_suite_CellTestSuite_test_formula : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_CellTestSuite_test_formula() : CxxTest::RealTestDescription( Tests_CellTestSuite, suiteDescription_CellTestSuite, 159, "test_formula" ) {}
 void runTest() { suite_CellTestSuite.test_formula(); }
} testDescription_suite_CellTestSuite_test_formula;

static class TestDescription_suite_CellTestSuite_test_boolean : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_CellTestSuite_test_boolean() : CxxTest::RealTestDescription( Tests_CellTestSuite, suiteDescription_CellTestSuite, 168, "test_boolean" ) {}
 void runTest() { suite_CellTestSuite.test_boolean(); }
} testDescription_suite_CellTestSuite_test_boolean;

static class TestDescription_suite_CellTestSuite_test_leading_zero : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_CellTestSuite_test_leading_zero() : CxxTest::RealTestDescription( Tests_CellTestSuite, suiteDescription_CellTestSuite, 179, "test_leading_zero" ) {}
 void runTest() { suite_CellTestSuite.test_leading_zero(); }
} testDescription_suite_CellTestSuite_test_leading_zero;

static class TestDescription_suite_CellTestSuite_test_error_codes : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_CellTestSuite_test_error_codes() : CxxTest::RealTestDescription( Tests_CellTestSuite, suiteDescription_CellTestSuite, 188, "test_error_codes" ) {}
 void runTest() { suite_CellTestSuite.test_error_codes(); }
} testDescription_suite_CellTestSuite_test_error_codes;

static class TestDescription_suite_CellTestSuite_test_data_type_check : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_CellTestSuite_test_data_type_check() : CxxTest::RealTestDescription( Tests_CellTestSuite, suiteDescription_CellTestSuite, 202, "test_data_type_check" ) {}
 void runTest() { suite_CellTestSuite.test_data_type_check(); }
} testDescription_suite_CellTestSuite_test_data_type_check;

static class TestDescription_suite_CellTestSuite_test_set_bad_type : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_CellTestSuite_test_set_bad_type() : CxxTest::RealTestDescription( Tests_CellTestSuite, suiteDescription_CellTestSuite, 220, "test_set_bad_type" ) {}
 void runTest() { suite_CellTestSuite.test_set_bad_type(); }
} testDescription_suite_CellTestSuite_test_set_bad_type;

static class TestDescription_suite_CellTestSuite_test_time : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_CellTestSuite_test_time() : CxxTest::RealTestDescription( Tests_CellTestSuite, suiteDescription_CellTestSuite, 230, "test_time" ) {}
 void runTest() { suite_CellTestSuite.test_time(); }
} testDescription_suite_CellTestSuite_test_time;

static class TestDescription_suite_CellTestSuite_test_date_format_on_non_date : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_CellTestSuite_test_date_format_on_non_date() : CxxTest::RealTestDescription( Tests_CellTestSuite, suiteDescription_CellTestSuite, 252, "test_date_format_on_non_date" ) {}
 void runTest() { suite_CellTestSuite.test_date_format_on_non_date(); }
} testDescription_suite_CellTestSuite_test_date_format_on_non_date;

static class TestDescription_suite_CellTestSuite_test_set_get_date : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_CellTestSuite_test_set_get_date() : CxxTest::RealTestDescription( Tests_CellTestSuite, suiteDescription_CellTestSuite, 265, "test_set_get_date" ) {}
 void runTest() { suite_CellTestSuite.test_set_get_date(); }
} testDescription_suite_CellTestSuite_test_set_get_date;

static class TestDescription_suite_CellTestSuite_test_repr : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_CellTestSuite_test_repr() : CxxTest::RealTestDescription( Tests_CellTestSuite, suiteDescription_CellTestSuite, 283, "test_repr" ) {}
 void runTest() { suite_CellTestSuite.test_repr(); }
} testDescription_suite_CellTestSuite_test_repr;

static class TestDescription_suite_CellTestSuite_test_is_date : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_CellTestSuite_test_is_date() : CxxTest::RealTestDescription( Tests_CellTestSuite, suiteDescription_CellTestSuite, 292, "test_is_date" ) {}
 void runTest() { suite_CellTestSuite.test_is_date(); }
} testDescription_suite_CellTestSuite_test_is_date;

static class TestDescription_suite_CellTestSuite_test_is_not_date_color_format : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_CellTestSuite_test_is_not_date_color_format() : CxxTest::RealTestDescription( Tests_CellTestSuite, suiteDescription_CellTestSuite, 309, "test_is_not_date_color_format" ) {}
 void runTest() { suite_CellTestSuite.test_is_not_date_color_format(); }
} testDescription_suite_CellTestSuite_test_is_not_date_color_format;

#include "/Users/thomas/Development/xlnt/source/tests/ChartTestSuite.h"

static ChartTestSuite suite_ChartTestSuite;

static CxxTest::List Tests_ChartTestSuite = { 0, 0 };
CxxTest::StaticSuiteDescription suiteDescription_ChartTestSuite( "../../source/tests/ChartTestSuite.h", 8, "ChartTestSuite", suite_ChartTestSuite, Tests_ChartTestSuite );

static class TestDescription_suite_ChartTestSuite_test_write_title : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ChartTestSuite_test_write_title() : CxxTest::RealTestDescription( Tests_ChartTestSuite, suiteDescription_ChartTestSuite, 36, "test_write_title" ) {}
 void runTest() { suite_ChartTestSuite.test_write_title(); }
} testDescription_suite_ChartTestSuite_test_write_title;

static class TestDescription_suite_ChartTestSuite_test_write_xaxis : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ChartTestSuite_test_write_xaxis() : CxxTest::RealTestDescription( Tests_ChartTestSuite, suiteDescription_ChartTestSuite, 44, "test_write_xaxis" ) {}
 void runTest() { suite_ChartTestSuite.test_write_xaxis(); }
} testDescription_suite_ChartTestSuite_test_write_xaxis;

static class TestDescription_suite_ChartTestSuite_test_write_yaxis : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ChartTestSuite_test_write_yaxis() : CxxTest::RealTestDescription( Tests_ChartTestSuite, suiteDescription_ChartTestSuite, 52, "test_write_yaxis" ) {}
 void runTest() { suite_ChartTestSuite.test_write_yaxis(); }
} testDescription_suite_ChartTestSuite_test_write_yaxis;

static class TestDescription_suite_ChartTestSuite_test_write_series : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ChartTestSuite_test_write_series() : CxxTest::RealTestDescription( Tests_ChartTestSuite, suiteDescription_ChartTestSuite, 60, "test_write_series" ) {}
 void runTest() { suite_ChartTestSuite.test_write_series(); }
} testDescription_suite_ChartTestSuite_test_write_series;

static class TestDescription_suite_ChartTestSuite_test_write_legend : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ChartTestSuite_test_write_legend() : CxxTest::RealTestDescription( Tests_ChartTestSuite, suiteDescription_ChartTestSuite, 66, "test_write_legend" ) {}
 void runTest() { suite_ChartTestSuite.test_write_legend(); }
} testDescription_suite_ChartTestSuite_test_write_legend;

static class TestDescription_suite_ChartTestSuite_test_no_write_legend : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ChartTestSuite_test_no_write_legend() : CxxTest::RealTestDescription( Tests_ChartTestSuite, suiteDescription_ChartTestSuite, 72, "test_no_write_legend" ) {}
 void runTest() { suite_ChartTestSuite.test_no_write_legend(); }
} testDescription_suite_ChartTestSuite_test_no_write_legend;

static class TestDescription_suite_ChartTestSuite_test_write_print_settings : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ChartTestSuite_test_write_print_settings() : CxxTest::RealTestDescription( Tests_ChartTestSuite, suiteDescription_ChartTestSuite, 92, "test_write_print_settings" ) {}
 void runTest() { suite_ChartTestSuite.test_write_print_settings(); }
} testDescription_suite_ChartTestSuite_test_write_print_settings;

static class TestDescription_suite_ChartTestSuite_test_write_chart : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ChartTestSuite_test_write_chart() : CxxTest::RealTestDescription( Tests_ChartTestSuite, suiteDescription_ChartTestSuite, 98, "test_write_chart" ) {}
 void runTest() { suite_ChartTestSuite.test_write_chart(); }
} testDescription_suite_ChartTestSuite_test_write_chart;

static class TestDescription_suite_ChartTestSuite_test_write_xaxis_scatter : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ChartTestSuite_test_write_xaxis_scatter() : CxxTest::RealTestDescription( Tests_ChartTestSuite, suiteDescription_ChartTestSuite, 123, "test_write_xaxis_scatter" ) {}
 void runTest() { suite_ChartTestSuite.test_write_xaxis_scatter(); }
} testDescription_suite_ChartTestSuite_test_write_xaxis_scatter;

static class TestDescription_suite_ChartTestSuite_test_write_yaxis_scatter : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ChartTestSuite_test_write_yaxis_scatter() : CxxTest::RealTestDescription( Tests_ChartTestSuite, suiteDescription_ChartTestSuite, 130, "test_write_yaxis_scatter" ) {}
 void runTest() { suite_ChartTestSuite.test_write_yaxis_scatter(); }
} testDescription_suite_ChartTestSuite_test_write_yaxis_scatter;

static class TestDescription_suite_ChartTestSuite_test_write_series_scatter : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ChartTestSuite_test_write_series_scatter() : CxxTest::RealTestDescription( Tests_ChartTestSuite, suiteDescription_ChartTestSuite, 137, "test_write_series_scatter" ) {}
 void runTest() { suite_ChartTestSuite.test_write_series_scatter(); }
} testDescription_suite_ChartTestSuite_test_write_series_scatter;

static class TestDescription_suite_ChartTestSuite_test_write_legend_scatter : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ChartTestSuite_test_write_legend_scatter() : CxxTest::RealTestDescription( Tests_ChartTestSuite, suiteDescription_ChartTestSuite, 143, "test_write_legend_scatter" ) {}
 void runTest() { suite_ChartTestSuite.test_write_legend_scatter(); }
} testDescription_suite_ChartTestSuite_test_write_legend_scatter;

static class TestDescription_suite_ChartTestSuite_test_write_print_settings_scatter : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ChartTestSuite_test_write_print_settings_scatter() : CxxTest::RealTestDescription( Tests_ChartTestSuite, suiteDescription_ChartTestSuite, 149, "test_write_print_settings_scatter" ) {}
 void runTest() { suite_ChartTestSuite.test_write_print_settings_scatter(); }
} testDescription_suite_ChartTestSuite_test_write_print_settings_scatter;

static class TestDescription_suite_ChartTestSuite_test_write_chart_scatter : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ChartTestSuite_test_write_chart_scatter() : CxxTest::RealTestDescription( Tests_ChartTestSuite, suiteDescription_ChartTestSuite, 155, "test_write_chart_scatter" ) {}
 void runTest() { suite_ChartTestSuite.test_write_chart_scatter(); }
} testDescription_suite_ChartTestSuite_test_write_chart_scatter;

#include "/Users/thomas/Development/xlnt/source/tests/DumpTestSuite.h"

static DumpTestSuite suite_DumpTestSuite;

static CxxTest::List Tests_DumpTestSuite = { 0, 0 };
CxxTest::StaticSuiteDescription suiteDescription_DumpTestSuite( "../../source/tests/DumpTestSuite.h", 9, "DumpTestSuite", suite_DumpTestSuite, Tests_DumpTestSuite );

static class TestDescription_suite_DumpTestSuite_test_dump_sheet_title : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_DumpTestSuite_test_dump_sheet_title() : CxxTest::RealTestDescription( Tests_DumpTestSuite, suiteDescription_DumpTestSuite, 19, "test_dump_sheet_title" ) {}
 void runTest() { suite_DumpTestSuite.test_dump_sheet_title(); }
} testDescription_suite_DumpTestSuite_test_dump_sheet_title;

static class TestDescription_suite_DumpTestSuite_test_dump_sheet : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_DumpTestSuite_test_dump_sheet() : CxxTest::RealTestDescription( Tests_DumpTestSuite, suiteDescription_DumpTestSuite, 34, "test_dump_sheet" ) {}
 void runTest() { suite_DumpTestSuite.test_dump_sheet(); }
} testDescription_suite_DumpTestSuite_test_dump_sheet;

static class TestDescription_suite_DumpTestSuite_test_table_builder : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_DumpTestSuite_test_table_builder() : CxxTest::RealTestDescription( Tests_DumpTestSuite, suiteDescription_DumpTestSuite, 101, "test_table_builder" ) {}
 void runTest() { suite_DumpTestSuite.test_table_builder(); }
} testDescription_suite_DumpTestSuite_test_table_builder;

static class TestDescription_suite_DumpTestSuite_test_open_too_many_files : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_DumpTestSuite_test_open_too_many_files() : CxxTest::RealTestDescription( Tests_DumpTestSuite, suiteDescription_DumpTestSuite, 123, "test_open_too_many_files" ) {}
 void runTest() { suite_DumpTestSuite.test_open_too_many_files(); }
} testDescription_suite_DumpTestSuite_test_open_too_many_files;

static class TestDescription_suite_DumpTestSuite_test_dump_twice : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_DumpTestSuite_test_dump_twice() : CxxTest::RealTestDescription( Tests_DumpTestSuite, suiteDescription_DumpTestSuite, 137, "test_dump_twice" ) {}
 void runTest() { suite_DumpTestSuite.test_dump_twice(); }
} testDescription_suite_DumpTestSuite_test_dump_twice;

static class TestDescription_suite_DumpTestSuite_test_append_after_save : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_DumpTestSuite_test_append_after_save() : CxxTest::RealTestDescription( Tests_DumpTestSuite, suiteDescription_DumpTestSuite, 152, "test_append_after_save" ) {}
 void runTest() { suite_DumpTestSuite.test_append_after_save(); }
} testDescription_suite_DumpTestSuite_test_append_after_save;

#include "/Users/thomas/Development/xlnt/source/tests/IterTestSuite.h"

static IterTestSuite suite_IterTestSuite;

static CxxTest::List Tests_IterTestSuite = { 0, 0 };
CxxTest::StaticSuiteDescription suiteDescription_IterTestSuite( "../../source/tests/IterTestSuite.h", 8, "IterTestSuite", suite_IterTestSuite, Tests_IterTestSuite );

static class TestDescription_suite_IterTestSuite_test_get_dimensions : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_IterTestSuite_test_get_dimensions() : CxxTest::RealTestDescription( Tests_IterTestSuite, suiteDescription_IterTestSuite, 16, "test_get_dimensions" ) {}
 void runTest() { suite_IterTestSuite.test_get_dimensions(); }
} testDescription_suite_IterTestSuite_test_get_dimensions;

static class TestDescription_suite_IterTestSuite_test_read_fast_integrated : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_IterTestSuite_test_read_fast_integrated() : CxxTest::RealTestDescription( Tests_IterTestSuite, suiteDescription_IterTestSuite, 32, "test_read_fast_integrated" ) {}
 void runTest() { suite_IterTestSuite.test_read_fast_integrated(); }
} testDescription_suite_IterTestSuite_test_read_fast_integrated;

static class TestDescription_suite_IterTestSuite_test_get_boundaries_range : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_IterTestSuite_test_get_boundaries_range() : CxxTest::RealTestDescription( Tests_IterTestSuite, suiteDescription_IterTestSuite, 52, "test_get_boundaries_range" ) {}
 void runTest() { suite_IterTestSuite.test_get_boundaries_range(); }
} testDescription_suite_IterTestSuite_test_get_boundaries_range;

static class TestDescription_suite_IterTestSuite_test_get_boundaries_one : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_IterTestSuite_test_get_boundaries_one() : CxxTest::RealTestDescription( Tests_IterTestSuite, suiteDescription_IterTestSuite, 57, "test_get_boundaries_one" ) {}
 void runTest() { suite_IterTestSuite.test_get_boundaries_one(); }
} testDescription_suite_IterTestSuite_test_get_boundaries_one;

static class TestDescription_suite_IterTestSuite_test_read_single_cell_range : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_IterTestSuite_test_read_single_cell_range() : CxxTest::RealTestDescription( Tests_IterTestSuite, suiteDescription_IterTestSuite, 62, "test_read_single_cell_range" ) {}
 void runTest() { suite_IterTestSuite.test_read_single_cell_range(); }
} testDescription_suite_IterTestSuite_test_read_single_cell_range;

static class TestDescription_suite_IterTestSuite_test_read_fast_integrated2 : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_IterTestSuite_test_read_fast_integrated2() : CxxTest::RealTestDescription( Tests_IterTestSuite, suiteDescription_IterTestSuite, 70, "test_read_fast_integrated2" ) {}
 void runTest() { suite_IterTestSuite.test_read_fast_integrated2(); }
} testDescription_suite_IterTestSuite_test_read_fast_integrated2;

static class TestDescription_suite_IterTestSuite_test_read_single_cell_date : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_IterTestSuite_test_read_single_cell_date() : CxxTest::RealTestDescription( Tests_IterTestSuite, suiteDescription_IterTestSuite, 88, "test_read_single_cell_date" ) {}
 void runTest() { suite_IterTestSuite.test_read_single_cell_date(); }
} testDescription_suite_IterTestSuite_test_read_single_cell_date;

#include "/Users/thomas/Development/xlnt/source/tests/NamedRangeTestSuite.h"

static NamedRangeTestSuite suite_NamedRangeTestSuite;

static CxxTest::List Tests_NamedRangeTestSuite = { 0, 0 };
CxxTest::StaticSuiteDescription suiteDescription_NamedRangeTestSuite( "../../source/tests/NamedRangeTestSuite.h", 8, "NamedRangeTestSuite", suite_NamedRangeTestSuite, Tests_NamedRangeTestSuite );

static class TestDescription_suite_NamedRangeTestSuite_test_split : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NamedRangeTestSuite_test_split() : CxxTest::RealTestDescription( Tests_NamedRangeTestSuite, suiteDescription_NamedRangeTestSuite, 16, "test_split" ) {}
 void runTest() { suite_NamedRangeTestSuite.test_split(); }
} testDescription_suite_NamedRangeTestSuite_test_split;

static class TestDescription_suite_NamedRangeTestSuite_test_split_no_quotes : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NamedRangeTestSuite_test_split_no_quotes() : CxxTest::RealTestDescription( Tests_NamedRangeTestSuite, suiteDescription_NamedRangeTestSuite, 21, "test_split_no_quotes" ) {}
 void runTest() { suite_NamedRangeTestSuite.test_split_no_quotes(); }
} testDescription_suite_NamedRangeTestSuite_test_split_no_quotes;

static class TestDescription_suite_NamedRangeTestSuite_test_bad_range_name : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NamedRangeTestSuite_test_bad_range_name() : CxxTest::RealTestDescription( Tests_NamedRangeTestSuite, suiteDescription_NamedRangeTestSuite, 26, "test_bad_range_name" ) {}
 void runTest() { suite_NamedRangeTestSuite.test_bad_range_name(); }
} testDescription_suite_NamedRangeTestSuite_test_bad_range_name;

static class TestDescription_suite_NamedRangeTestSuite_test_range_name_worksheet_special_chars : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NamedRangeTestSuite_test_range_name_worksheet_special_chars() : CxxTest::RealTestDescription( Tests_NamedRangeTestSuite, suiteDescription_NamedRangeTestSuite, 31, "test_range_name_worksheet_special_chars" ) {}
 void runTest() { suite_NamedRangeTestSuite.test_range_name_worksheet_special_chars(); }
} testDescription_suite_NamedRangeTestSuite_test_range_name_worksheet_special_chars;

static class TestDescription_suite_NamedRangeTestSuite_test_read_named_ranges : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NamedRangeTestSuite_test_read_named_ranges() : CxxTest::RealTestDescription( Tests_NamedRangeTestSuite, suiteDescription_NamedRangeTestSuite, 70, "test_read_named_ranges" ) {}
 void runTest() { suite_NamedRangeTestSuite.test_read_named_ranges(); }
} testDescription_suite_NamedRangeTestSuite_test_read_named_ranges;

static class TestDescription_suite_NamedRangeTestSuite_test_oddly_shaped_named_ranges : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NamedRangeTestSuite_test_oddly_shaped_named_ranges() : CxxTest::RealTestDescription( Tests_NamedRangeTestSuite, suiteDescription_NamedRangeTestSuite, 99, "test_oddly_shaped_named_ranges" ) {}
 void runTest() { suite_NamedRangeTestSuite.test_oddly_shaped_named_ranges(); }
} testDescription_suite_NamedRangeTestSuite_test_oddly_shaped_named_ranges;

static class TestDescription_suite_NamedRangeTestSuite_test_merged_cells_named_range : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NamedRangeTestSuite_test_merged_cells_named_range() : CxxTest::RealTestDescription( Tests_NamedRangeTestSuite, suiteDescription_NamedRangeTestSuite, 122, "test_merged_cells_named_range" ) {}
 void runTest() { suite_NamedRangeTestSuite.test_merged_cells_named_range(); }
} testDescription_suite_NamedRangeTestSuite_test_merged_cells_named_range;

static class TestDescription_suite_NamedRangeTestSuite_test_has_ranges : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NamedRangeTestSuite_test_has_ranges() : CxxTest::RealTestDescription( Tests_NamedRangeTestSuite, suiteDescription_NamedRangeTestSuite, 148, "test_has_ranges" ) {}
 void runTest() { suite_NamedRangeTestSuite.test_has_ranges(); }
} testDescription_suite_NamedRangeTestSuite_test_has_ranges;

static class TestDescription_suite_NamedRangeTestSuite_test_workbook_has_normal_range : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NamedRangeTestSuite_test_workbook_has_normal_range() : CxxTest::RealTestDescription( Tests_NamedRangeTestSuite, suiteDescription_NamedRangeTestSuite, 154, "test_workbook_has_normal_range" ) {}
 void runTest() { suite_NamedRangeTestSuite.test_workbook_has_normal_range(); }
} testDescription_suite_NamedRangeTestSuite_test_workbook_has_normal_range;

static class TestDescription_suite_NamedRangeTestSuite_test_workbook_has_value_range : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NamedRangeTestSuite_test_workbook_has_value_range() : CxxTest::RealTestDescription( Tests_NamedRangeTestSuite, suiteDescription_NamedRangeTestSuite, 160, "test_workbook_has_value_range" ) {}
 void runTest() { suite_NamedRangeTestSuite.test_workbook_has_value_range(); }
} testDescription_suite_NamedRangeTestSuite_test_workbook_has_value_range;

static class TestDescription_suite_NamedRangeTestSuite_test_worksheet_range : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NamedRangeTestSuite_test_worksheet_range() : CxxTest::RealTestDescription( Tests_NamedRangeTestSuite, suiteDescription_NamedRangeTestSuite, 167, "test_worksheet_range" ) {}
 void runTest() { suite_NamedRangeTestSuite.test_worksheet_range(); }
} testDescription_suite_NamedRangeTestSuite_test_worksheet_range;

static class TestDescription_suite_NamedRangeTestSuite_test_worksheet_range_error_on_value_range : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NamedRangeTestSuite_test_worksheet_range_error_on_value_range() : CxxTest::RealTestDescription( Tests_NamedRangeTestSuite, suiteDescription_NamedRangeTestSuite, 172, "test_worksheet_range_error_on_value_range" ) {}
 void runTest() { suite_NamedRangeTestSuite.test_worksheet_range_error_on_value_range(); }
} testDescription_suite_NamedRangeTestSuite_test_worksheet_range_error_on_value_range;

static class TestDescription_suite_NamedRangeTestSuite_test_handles_scope : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NamedRangeTestSuite_test_handles_scope() : CxxTest::RealTestDescription( Tests_NamedRangeTestSuite, suiteDescription_NamedRangeTestSuite, 199, "test_handles_scope" ) {}
 void runTest() { suite_NamedRangeTestSuite.test_handles_scope(); }
} testDescription_suite_NamedRangeTestSuite_test_handles_scope;

static class TestDescription_suite_NamedRangeTestSuite_test_can_be_saved : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NamedRangeTestSuite_test_can_be_saved() : CxxTest::RealTestDescription( Tests_NamedRangeTestSuite, suiteDescription_NamedRangeTestSuite, 206, "test_can_be_saved" ) {}
 void runTest() { suite_NamedRangeTestSuite.test_can_be_saved(); }
} testDescription_suite_NamedRangeTestSuite_test_can_be_saved;

#include "/Users/thomas/Development/xlnt/source/tests/NullableTestSuite.h"

static NullableTestSuite suite_NullableTestSuite;

static CxxTest::List Tests_NullableTestSuite = { 0, 0 };
CxxTest::StaticSuiteDescription suiteDescription_NullableTestSuite( "../../source/tests/NullableTestSuite.h", 8, "NullableTestSuite", suite_NullableTestSuite, Tests_NullableTestSuite );

static class TestDescription_suite_NullableTestSuite_test_has_value : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NullableTestSuite_test_has_value() : CxxTest::RealTestDescription( Tests_NullableTestSuite, suiteDescription_NullableTestSuite, 16, "test_has_value" ) {}
 void runTest() { suite_NullableTestSuite.test_has_value(); }
} testDescription_suite_NullableTestSuite_test_has_value;

static class TestDescription_suite_NullableTestSuite_test_get_value : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NullableTestSuite_test_get_value() : CxxTest::RealTestDescription( Tests_NullableTestSuite, suiteDescription_NullableTestSuite, 22, "test_get_value" ) {}
 void runTest() { suite_NullableTestSuite.test_get_value(); }
} testDescription_suite_NullableTestSuite_test_get_value;

static class TestDescription_suite_NullableTestSuite_test_equality : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NullableTestSuite_test_equality() : CxxTest::RealTestDescription( Tests_NullableTestSuite, suiteDescription_NullableTestSuite, 29, "test_equality" ) {}
 void runTest() { suite_NullableTestSuite.test_equality(); }
} testDescription_suite_NullableTestSuite_test_equality;

static class TestDescription_suite_NullableTestSuite_test_assignment : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NullableTestSuite_test_assignment() : CxxTest::RealTestDescription( Tests_NullableTestSuite, suiteDescription_NullableTestSuite, 36, "test_assignment" ) {}
 void runTest() { suite_NullableTestSuite.test_assignment(); }
} testDescription_suite_NullableTestSuite_test_assignment;

static class TestDescription_suite_NullableTestSuite_test_copy_constructor : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NullableTestSuite_test_copy_constructor() : CxxTest::RealTestDescription( Tests_NullableTestSuite, suiteDescription_NullableTestSuite, 44, "test_copy_constructor" ) {}
 void runTest() { suite_NullableTestSuite.test_copy_constructor(); }
} testDescription_suite_NullableTestSuite_test_copy_constructor;

#include "/Users/thomas/Development/xlnt/source/tests/NumberFormatTestSuite.h"

static NumberFormatTestSuite suite_NumberFormatTestSuite;

static CxxTest::List Tests_NumberFormatTestSuite = { 0, 0 };
CxxTest::StaticSuiteDescription suiteDescription_NumberFormatTestSuite( "../../source/tests/NumberFormatTestSuite.h", 8, "NumberFormatTestSuite", suite_NumberFormatTestSuite, Tests_NumberFormatTestSuite );

static class TestDescription_suite_NumberFormatTestSuite_test_convert_date_to_julian : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NumberFormatTestSuite_test_convert_date_to_julian() : CxxTest::RealTestDescription( Tests_NumberFormatTestSuite, suiteDescription_NumberFormatTestSuite, 23, "test_convert_date_to_julian" ) {}
 void runTest() { suite_NumberFormatTestSuite.test_convert_date_to_julian(); }
} testDescription_suite_NumberFormatTestSuite_test_convert_date_to_julian;

static class TestDescription_suite_NumberFormatTestSuite_test_convert_date_from_julian : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NumberFormatTestSuite_test_convert_date_from_julian() : CxxTest::RealTestDescription( Tests_NumberFormatTestSuite, suiteDescription_NumberFormatTestSuite, 28, "test_convert_date_from_julian" ) {}
 void runTest() { suite_NumberFormatTestSuite.test_convert_date_from_julian(); }
} testDescription_suite_NumberFormatTestSuite_test_convert_date_from_julian;

static class TestDescription_suite_NumberFormatTestSuite_test_convert_datetime_to_julian : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NumberFormatTestSuite_test_convert_datetime_to_julian() : CxxTest::RealTestDescription( Tests_NumberFormatTestSuite, suiteDescription_NumberFormatTestSuite, 47, "test_convert_datetime_to_julian" ) {}
 void runTest() { suite_NumberFormatTestSuite.test_convert_datetime_to_julian(); }
} testDescription_suite_NumberFormatTestSuite_test_convert_datetime_to_julian;

static class TestDescription_suite_NumberFormatTestSuite_test_insert_float : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NumberFormatTestSuite_test_insert_float() : CxxTest::RealTestDescription( Tests_NumberFormatTestSuite, suiteDescription_NumberFormatTestSuite, 53, "test_insert_float" ) {}
 void runTest() { suite_NumberFormatTestSuite.test_insert_float(); }
} testDescription_suite_NumberFormatTestSuite_test_insert_float;

static class TestDescription_suite_NumberFormatTestSuite_test_insert_percentage : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NumberFormatTestSuite_test_insert_percentage() : CxxTest::RealTestDescription( Tests_NumberFormatTestSuite, suiteDescription_NumberFormatTestSuite, 59, "test_insert_percentage" ) {}
 void runTest() { suite_NumberFormatTestSuite.test_insert_percentage(); }
} testDescription_suite_NumberFormatTestSuite_test_insert_percentage;

static class TestDescription_suite_NumberFormatTestSuite_test_insert_datetime : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NumberFormatTestSuite_test_insert_datetime() : CxxTest::RealTestDescription( Tests_NumberFormatTestSuite, suiteDescription_NumberFormatTestSuite, 66, "test_insert_datetime" ) {}
 void runTest() { suite_NumberFormatTestSuite.test_insert_datetime(); }
} testDescription_suite_NumberFormatTestSuite_test_insert_datetime;

static class TestDescription_suite_NumberFormatTestSuite_test_insert_date : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NumberFormatTestSuite_test_insert_date() : CxxTest::RealTestDescription( Tests_NumberFormatTestSuite, suiteDescription_NumberFormatTestSuite, 72, "test_insert_date" ) {}
 void runTest() { suite_NumberFormatTestSuite.test_insert_date(); }
} testDescription_suite_NumberFormatTestSuite_test_insert_date;

static class TestDescription_suite_NumberFormatTestSuite_test_internal_date : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NumberFormatTestSuite_test_internal_date() : CxxTest::RealTestDescription( Tests_NumberFormatTestSuite, suiteDescription_NumberFormatTestSuite, 78, "test_internal_date" ) {}
 void runTest() { suite_NumberFormatTestSuite.test_internal_date(); }
} testDescription_suite_NumberFormatTestSuite_test_internal_date;

static class TestDescription_suite_NumberFormatTestSuite_test_datetime_interpretation : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NumberFormatTestSuite_test_datetime_interpretation() : CxxTest::RealTestDescription( Tests_NumberFormatTestSuite, suiteDescription_NumberFormatTestSuite, 85, "test_datetime_interpretation" ) {}
 void runTest() { suite_NumberFormatTestSuite.test_datetime_interpretation(); }
} testDescription_suite_NumberFormatTestSuite_test_datetime_interpretation;

static class TestDescription_suite_NumberFormatTestSuite_test_date_interpretation : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NumberFormatTestSuite_test_date_interpretation() : CxxTest::RealTestDescription( Tests_NumberFormatTestSuite, suiteDescription_NumberFormatTestSuite, 92, "test_date_interpretation" ) {}
 void runTest() { suite_NumberFormatTestSuite.test_date_interpretation(); }
} testDescription_suite_NumberFormatTestSuite_test_date_interpretation;

static class TestDescription_suite_NumberFormatTestSuite_test_number_format_style : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NumberFormatTestSuite_test_number_format_style() : CxxTest::RealTestDescription( Tests_NumberFormatTestSuite, suiteDescription_NumberFormatTestSuite, 99, "test_number_format_style" ) {}
 void runTest() { suite_NumberFormatTestSuite.test_number_format_style(); }
} testDescription_suite_NumberFormatTestSuite_test_number_format_style;

static class TestDescription_suite_NumberFormatTestSuite_test_date_format_on_non_date : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NumberFormatTestSuite_test_date_format_on_non_date() : CxxTest::RealTestDescription( Tests_NumberFormatTestSuite, suiteDescription_NumberFormatTestSuite, 106, "test_date_format_on_non_date" ) {}
 void runTest() { suite_NumberFormatTestSuite.test_date_format_on_non_date(); }
} testDescription_suite_NumberFormatTestSuite_test_date_format_on_non_date;

static class TestDescription_suite_NumberFormatTestSuite_test_1900_leap_year : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NumberFormatTestSuite_test_1900_leap_year() : CxxTest::RealTestDescription( Tests_NumberFormatTestSuite, suiteDescription_NumberFormatTestSuite, 128, "test_1900_leap_year" ) {}
 void runTest() { suite_NumberFormatTestSuite.test_1900_leap_year(); }
} testDescription_suite_NumberFormatTestSuite_test_1900_leap_year;

static class TestDescription_suite_NumberFormatTestSuite_test_bad_date : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NumberFormatTestSuite_test_bad_date() : CxxTest::RealTestDescription( Tests_NumberFormatTestSuite, suiteDescription_NumberFormatTestSuite, 134, "test_bad_date" ) {}
 void runTest() { suite_NumberFormatTestSuite.test_bad_date(); }
} testDescription_suite_NumberFormatTestSuite_test_bad_date;

static class TestDescription_suite_NumberFormatTestSuite_test_bad_julian_date : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NumberFormatTestSuite_test_bad_julian_date() : CxxTest::RealTestDescription( Tests_NumberFormatTestSuite, suiteDescription_NumberFormatTestSuite, 148, "test_bad_julian_date" ) {}
 void runTest() { suite_NumberFormatTestSuite.test_bad_julian_date(); }
} testDescription_suite_NumberFormatTestSuite_test_bad_julian_date;

static class TestDescription_suite_NumberFormatTestSuite_test_mac_date : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_NumberFormatTestSuite_test_mac_date() : CxxTest::RealTestDescription( Tests_NumberFormatTestSuite, suiteDescription_NumberFormatTestSuite, 153, "test_mac_date" ) {}
 void runTest() { suite_NumberFormatTestSuite.test_mac_date(); }
} testDescription_suite_NumberFormatTestSuite_test_mac_date;

#include "/Users/thomas/Development/xlnt/source/tests/PasswordHashTestSuite.h"

static PasswordHashTestSuite suite_PasswordHashTestSuite;

static CxxTest::List Tests_PasswordHashTestSuite = { 0, 0 };
CxxTest::StaticSuiteDescription suiteDescription_PasswordHashTestSuite( "../../source/tests/PasswordHashTestSuite.h", 8, "PasswordHashTestSuite", suite_PasswordHashTestSuite, Tests_PasswordHashTestSuite );

static class TestDescription_suite_PasswordHashTestSuite_test_hasher : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_PasswordHashTestSuite_test_hasher() : CxxTest::RealTestDescription( Tests_PasswordHashTestSuite, suiteDescription_PasswordHashTestSuite, 16, "test_hasher" ) {}
 void runTest() { suite_PasswordHashTestSuite.test_hasher(); }
} testDescription_suite_PasswordHashTestSuite_test_hasher;

static class TestDescription_suite_PasswordHashTestSuite_test_sheet_protection : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_PasswordHashTestSuite_test_sheet_protection() : CxxTest::RealTestDescription( Tests_PasswordHashTestSuite, suiteDescription_PasswordHashTestSuite, 21, "test_sheet_protection" ) {}
 void runTest() { suite_PasswordHashTestSuite.test_sheet_protection(); }
} testDescription_suite_PasswordHashTestSuite_test_sheet_protection;

#include "/Users/thomas/Development/xlnt/source/tests/PropsTestSuite.h"

static PropsTestSuite suite_PropsTestSuite;

static CxxTest::List Tests_PropsTestSuite = { 0, 0 };
CxxTest::StaticSuiteDescription suiteDescription_PropsTestSuite( "../../source/tests/PropsTestSuite.h", 8, "PropsTestSuite", suite_PropsTestSuite, Tests_PropsTestSuite );

static class TestDescription_suite_PropsTestSuite_test_1 : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_PropsTestSuite_test_1() : CxxTest::RealTestDescription( Tests_PropsTestSuite, suiteDescription_PropsTestSuite, 16, "test_1" ) {}
 void runTest() { suite_PropsTestSuite.test_1(); }
} testDescription_suite_PropsTestSuite_test_1;

static class TestDescription_suite_PropsTestSuite_test_read_properties_core : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_PropsTestSuite_test_read_properties_core() : CxxTest::RealTestDescription( Tests_PropsTestSuite, suiteDescription_PropsTestSuite, 35, "test_read_properties_core" ) {}
 void runTest() { suite_PropsTestSuite.test_read_properties_core(); }
} testDescription_suite_PropsTestSuite_test_read_properties_core;

static class TestDescription_suite_PropsTestSuite_test_read_sheets_titles : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_PropsTestSuite_test_read_sheets_titles() : CxxTest::RealTestDescription( Tests_PropsTestSuite, suiteDescription_PropsTestSuite, 46, "test_read_sheets_titles" ) {}
 void runTest() { suite_PropsTestSuite.test_read_sheets_titles(); }
} testDescription_suite_PropsTestSuite_test_read_sheets_titles;

static class TestDescription_suite_PropsTestSuite_test_read_properties_core2 : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_PropsTestSuite_test_read_properties_core2() : CxxTest::RealTestDescription( Tests_PropsTestSuite, suiteDescription_PropsTestSuite, 67, "test_read_properties_core2" ) {}
 void runTest() { suite_PropsTestSuite.test_read_properties_core2(); }
} testDescription_suite_PropsTestSuite_test_read_properties_core2;

static class TestDescription_suite_PropsTestSuite_test_read_sheets_titles2 : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_PropsTestSuite_test_read_sheets_titles2() : CxxTest::RealTestDescription( Tests_PropsTestSuite, suiteDescription_PropsTestSuite, 74, "test_read_sheets_titles2" ) {}
 void runTest() { suite_PropsTestSuite.test_read_sheets_titles2(); }
} testDescription_suite_PropsTestSuite_test_read_sheets_titles2;

static class TestDescription_suite_PropsTestSuite_test_write_properties_core : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_PropsTestSuite_test_write_properties_core() : CxxTest::RealTestDescription( Tests_PropsTestSuite, suiteDescription_PropsTestSuite, 82, "test_write_properties_core" ) {}
 void runTest() { suite_PropsTestSuite.test_write_properties_core(); }
} testDescription_suite_PropsTestSuite_test_write_properties_core;

static class TestDescription_suite_PropsTestSuite_test_write_properties_app : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_PropsTestSuite_test_write_properties_app() : CxxTest::RealTestDescription( Tests_PropsTestSuite, suiteDescription_PropsTestSuite, 94, "test_write_properties_app" ) {}
 void runTest() { suite_PropsTestSuite.test_write_properties_app(); }
} testDescription_suite_PropsTestSuite_test_write_properties_app;

#include "/Users/thomas/Development/xlnt/source/tests/ReadTestSuite.h"

static ReadTestSuite suite_ReadTestSuite;

static CxxTest::List Tests_ReadTestSuite = { 0, 0 };
CxxTest::StaticSuiteDescription suiteDescription_ReadTestSuite( "../../source/tests/ReadTestSuite.h", 8, "ReadTestSuite", suite_ReadTestSuite, Tests_ReadTestSuite );

static class TestDescription_suite_ReadTestSuite_test_read_standalone_worksheet : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ReadTestSuite_test_read_standalone_worksheet() : CxxTest::RealTestDescription( Tests_ReadTestSuite, suiteDescription_ReadTestSuite, 16, "test_read_standalone_worksheet" ) {}
 void runTest() { suite_ReadTestSuite.test_read_standalone_worksheet(); }
} testDescription_suite_ReadTestSuite_test_read_standalone_worksheet;

static class TestDescription_suite_ReadTestSuite_test_read_standard_workbook : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ReadTestSuite_test_read_standard_workbook() : CxxTest::RealTestDescription( Tests_ReadTestSuite, suiteDescription_ReadTestSuite, 32, "test_read_standard_workbook" ) {}
 void runTest() { suite_ReadTestSuite.test_read_standard_workbook(); }
} testDescription_suite_ReadTestSuite_test_read_standard_workbook;

static class TestDescription_suite_ReadTestSuite_test_read_standard_workbook_from_fileobj : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ReadTestSuite_test_read_standard_workbook_from_fileobj() : CxxTest::RealTestDescription( Tests_ReadTestSuite, suiteDescription_ReadTestSuite, 39, "test_read_standard_workbook_from_fileobj" ) {}
 void runTest() { suite_ReadTestSuite.test_read_standard_workbook_from_fileobj(); }
} testDescription_suite_ReadTestSuite_test_read_standard_workbook_from_fileobj;

static class TestDescription_suite_ReadTestSuite_test_read_worksheet : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ReadTestSuite_test_read_worksheet() : CxxTest::RealTestDescription( Tests_ReadTestSuite, suiteDescription_ReadTestSuite, 47, "test_read_worksheet" ) {}
 void runTest() { suite_ReadTestSuite.test_read_worksheet(); }
} testDescription_suite_ReadTestSuite_test_read_worksheet;

static class TestDescription_suite_ReadTestSuite_test_read_nostring_workbook : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ReadTestSuite_test_read_nostring_workbook() : CxxTest::RealTestDescription( Tests_ReadTestSuite, suiteDescription_ReadTestSuite, 57, "test_read_nostring_workbook" ) {}
 void runTest() { suite_ReadTestSuite.test_read_nostring_workbook(); }
} testDescription_suite_ReadTestSuite_test_read_nostring_workbook;

static class TestDescription_suite_ReadTestSuite_test_read_empty_file : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ReadTestSuite_test_read_empty_file() : CxxTest::RealTestDescription( Tests_ReadTestSuite, suiteDescription_ReadTestSuite, 64, "test_read_empty_file" ) {}
 void runTest() { suite_ReadTestSuite.test_read_empty_file(); }
} testDescription_suite_ReadTestSuite_test_read_empty_file;

static class TestDescription_suite_ReadTestSuite_test_read_empty_archive : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ReadTestSuite_test_read_empty_archive() : CxxTest::RealTestDescription( Tests_ReadTestSuite, suiteDescription_ReadTestSuite, 72, "test_read_empty_archive" ) {}
 void runTest() { suite_ReadTestSuite.test_read_empty_archive(); }
} testDescription_suite_ReadTestSuite_test_read_empty_archive;

static class TestDescription_suite_ReadTestSuite_test_read_dimension : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ReadTestSuite_test_read_dimension() : CxxTest::RealTestDescription( Tests_ReadTestSuite, suiteDescription_ReadTestSuite, 78, "test_read_dimension" ) {}
 void runTest() { suite_ReadTestSuite.test_read_dimension(); }
} testDescription_suite_ReadTestSuite_test_read_dimension;

static class TestDescription_suite_ReadTestSuite_test_calculate_dimension_iter : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ReadTestSuite_test_calculate_dimension_iter() : CxxTest::RealTestDescription( Tests_ReadTestSuite, suiteDescription_ReadTestSuite, 92, "test_calculate_dimension_iter" ) {}
 void runTest() { suite_ReadTestSuite.test_calculate_dimension_iter(); }
} testDescription_suite_ReadTestSuite_test_calculate_dimension_iter;

static class TestDescription_suite_ReadTestSuite_test_get_highest_row_iter : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ReadTestSuite_test_get_highest_row_iter() : CxxTest::RealTestDescription( Tests_ReadTestSuite, suiteDescription_ReadTestSuite, 101, "test_get_highest_row_iter" ) {}
 void runTest() { suite_ReadTestSuite.test_get_highest_row_iter(); }
} testDescription_suite_ReadTestSuite_test_get_highest_row_iter;

static class TestDescription_suite_ReadTestSuite_test_read_workbook_with_no_properties : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ReadTestSuite_test_read_workbook_with_no_properties() : CxxTest::RealTestDescription( Tests_ReadTestSuite, suiteDescription_ReadTestSuite, 110, "test_read_workbook_with_no_properties" ) {}
 void runTest() { suite_ReadTestSuite.test_read_workbook_with_no_properties(); }
} testDescription_suite_ReadTestSuite_test_read_workbook_with_no_properties;

static class TestDescription_suite_ReadTestSuite_test_read_general_style : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ReadTestSuite_test_read_general_style() : CxxTest::RealTestDescription( Tests_ReadTestSuite, suiteDescription_ReadTestSuite, 125, "test_read_general_style" ) {}
 void runTest() { suite_ReadTestSuite.test_read_general_style(); }
} testDescription_suite_ReadTestSuite_test_read_general_style;

static class TestDescription_suite_ReadTestSuite_test_read_date_style : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ReadTestSuite_test_read_date_style() : CxxTest::RealTestDescription( Tests_ReadTestSuite, suiteDescription_ReadTestSuite, 131, "test_read_date_style" ) {}
 void runTest() { suite_ReadTestSuite.test_read_date_style(); }
} testDescription_suite_ReadTestSuite_test_read_date_style;

static class TestDescription_suite_ReadTestSuite_test_read_number_style : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ReadTestSuite_test_read_number_style() : CxxTest::RealTestDescription( Tests_ReadTestSuite, suiteDescription_ReadTestSuite, 137, "test_read_number_style" ) {}
 void runTest() { suite_ReadTestSuite.test_read_number_style(); }
} testDescription_suite_ReadTestSuite_test_read_number_style;

static class TestDescription_suite_ReadTestSuite_test_read_time_style : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ReadTestSuite_test_read_time_style() : CxxTest::RealTestDescription( Tests_ReadTestSuite, suiteDescription_ReadTestSuite, 143, "test_read_time_style" ) {}
 void runTest() { suite_ReadTestSuite.test_read_time_style(); }
} testDescription_suite_ReadTestSuite_test_read_time_style;

static class TestDescription_suite_ReadTestSuite_test_read_percentage_style : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ReadTestSuite_test_read_percentage_style() : CxxTest::RealTestDescription( Tests_ReadTestSuite, suiteDescription_ReadTestSuite, 149, "test_read_percentage_style" ) {}
 void runTest() { suite_ReadTestSuite.test_read_percentage_style(); }
} testDescription_suite_ReadTestSuite_test_read_percentage_style;

static class TestDescription_suite_ReadTestSuite_test_read_win_base_date : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ReadTestSuite_test_read_win_base_date() : CxxTest::RealTestDescription( Tests_ReadTestSuite, suiteDescription_ReadTestSuite, 166, "test_read_win_base_date" ) {}
 void runTest() { suite_ReadTestSuite.test_read_win_base_date(); }
} testDescription_suite_ReadTestSuite_test_read_win_base_date;

static class TestDescription_suite_ReadTestSuite_test_read_mac_base_date : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ReadTestSuite_test_read_mac_base_date() : CxxTest::RealTestDescription( Tests_ReadTestSuite, suiteDescription_ReadTestSuite, 171, "test_read_mac_base_date" ) {}
 void runTest() { suite_ReadTestSuite.test_read_mac_base_date(); }
} testDescription_suite_ReadTestSuite_test_read_mac_base_date;

static class TestDescription_suite_ReadTestSuite_test_read_date_style_mac : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ReadTestSuite_test_read_date_style_mac() : CxxTest::RealTestDescription( Tests_ReadTestSuite, suiteDescription_ReadTestSuite, 176, "test_read_date_style_mac" ) {}
 void runTest() { suite_ReadTestSuite.test_read_date_style_mac(); }
} testDescription_suite_ReadTestSuite_test_read_date_style_mac;

static class TestDescription_suite_ReadTestSuite_test_read_date_style_win : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ReadTestSuite_test_read_date_style_win() : CxxTest::RealTestDescription( Tests_ReadTestSuite, suiteDescription_ReadTestSuite, 182, "test_read_date_style_win" ) {}
 void runTest() { suite_ReadTestSuite.test_read_date_style_win(); }
} testDescription_suite_ReadTestSuite_test_read_date_style_win;

static class TestDescription_suite_ReadTestSuite_test_read_date_value : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ReadTestSuite_test_read_date_value() : CxxTest::RealTestDescription( Tests_ReadTestSuite, suiteDescription_ReadTestSuite, 188, "test_read_date_value" ) {}
 void runTest() { suite_ReadTestSuite.test_read_date_value(); }
} testDescription_suite_ReadTestSuite_test_read_date_value;

#include "/Users/thomas/Development/xlnt/source/tests/StringsTestSuite.h"

static StringsTestSuite suite_StringsTestSuite;

static CxxTest::List Tests_StringsTestSuite = { 0, 0 };
CxxTest::StaticSuiteDescription suiteDescription_StringsTestSuite( "../../source/tests/StringsTestSuite.h", 8, "StringsTestSuite", suite_StringsTestSuite, Tests_StringsTestSuite );

static class TestDescription_suite_StringsTestSuite_test_create_string_table : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_StringsTestSuite_test_create_string_table() : CxxTest::RealTestDescription( Tests_StringsTestSuite, suiteDescription_StringsTestSuite, 16, "test_create_string_table" ) {}
 void runTest() { suite_StringsTestSuite.test_create_string_table(); }
} testDescription_suite_StringsTestSuite_test_create_string_table;

static class TestDescription_suite_StringsTestSuite_test_read_string_table : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_StringsTestSuite_test_read_string_table() : CxxTest::RealTestDescription( Tests_StringsTestSuite, suiteDescription_StringsTestSuite, 27, "test_read_string_table" ) {}
 void runTest() { suite_StringsTestSuite.test_read_string_table(); }
} testDescription_suite_StringsTestSuite_test_read_string_table;

static class TestDescription_suite_StringsTestSuite_test_empty_string : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_StringsTestSuite_test_empty_string() : CxxTest::RealTestDescription( Tests_StringsTestSuite, suiteDescription_StringsTestSuite, 38, "test_empty_string" ) {}
 void runTest() { suite_StringsTestSuite.test_empty_string(); }
} testDescription_suite_StringsTestSuite_test_empty_string;

static class TestDescription_suite_StringsTestSuite_test_formatted_string_table : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_StringsTestSuite_test_formatted_string_table() : CxxTest::RealTestDescription( Tests_StringsTestSuite, suiteDescription_StringsTestSuite, 49, "test_formatted_string_table" ) {}
 void runTest() { suite_StringsTestSuite.test_formatted_string_table(); }
} testDescription_suite_StringsTestSuite_test_formatted_string_table;

#include "/Users/thomas/Development/xlnt/source/tests/StyleTestSuite.h"

static StyleTestSuite suite_StyleTestSuite;

static CxxTest::List Tests_StyleTestSuite = { 0, 0 };
CxxTest::StaticSuiteDescription suiteDescription_StyleTestSuite( "../../source/tests/StyleTestSuite.h", 8, "StyleTestSuite", suite_StyleTestSuite, Tests_StyleTestSuite );

static class TestDescription_suite_StyleTestSuite_test_create_style_table : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_StyleTestSuite_test_create_style_table() : CxxTest::RealTestDescription( Tests_StyleTestSuite, suiteDescription_StyleTestSuite, 29, "test_create_style_table" ) {}
 void runTest() { suite_StyleTestSuite.test_create_style_table(); }
} testDescription_suite_StyleTestSuite_test_create_style_table;

static class TestDescription_suite_StyleTestSuite_test_write_style_table : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_StyleTestSuite_test_write_style_table() : CxxTest::RealTestDescription( Tests_StyleTestSuite, suiteDescription_StyleTestSuite, 34, "test_write_style_table" ) {}
 void runTest() { suite_StyleTestSuite.test_write_style_table(); }
} testDescription_suite_StyleTestSuite_test_write_style_table;

static class TestDescription_suite_StyleTestSuite_test_no_style : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_StyleTestSuite_test_no_style() : CxxTest::RealTestDescription( Tests_StyleTestSuite, suiteDescription_StyleTestSuite, 46, "test_no_style" ) {}
 void runTest() { suite_StyleTestSuite.test_no_style(); }
} testDescription_suite_StyleTestSuite_test_no_style;

static class TestDescription_suite_StyleTestSuite_test_nb_style : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_StyleTestSuite_test_nb_style() : CxxTest::RealTestDescription( Tests_StyleTestSuite, suiteDescription_StyleTestSuite, 53, "test_nb_style" ) {}
 void runTest() { suite_StyleTestSuite.test_nb_style(); }
} testDescription_suite_StyleTestSuite_test_nb_style;

static class TestDescription_suite_StyleTestSuite_test_style_unicity : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_StyleTestSuite_test_style_unicity() : CxxTest::RealTestDescription( Tests_StyleTestSuite, suiteDescription_StyleTestSuite, 67, "test_style_unicity" ) {}
 void runTest() { suite_StyleTestSuite.test_style_unicity(); }
} testDescription_suite_StyleTestSuite_test_style_unicity;

static class TestDescription_suite_StyleTestSuite_test_fonts : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_StyleTestSuite_test_fonts() : CxxTest::RealTestDescription( Tests_StyleTestSuite, suiteDescription_StyleTestSuite, 77, "test_fonts" ) {}
 void runTest() { suite_StyleTestSuite.test_fonts(); }
} testDescription_suite_StyleTestSuite_test_fonts;

static class TestDescription_suite_StyleTestSuite_test_fills : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_StyleTestSuite_test_fills() : CxxTest::RealTestDescription( Tests_StyleTestSuite, suiteDescription_StyleTestSuite, 91, "test_fills" ) {}
 void runTest() { suite_StyleTestSuite.test_fills(); }
} testDescription_suite_StyleTestSuite_test_fills;

static class TestDescription_suite_StyleTestSuite_test_borders : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_StyleTestSuite_test_borders() : CxxTest::RealTestDescription( Tests_StyleTestSuite, suiteDescription_StyleTestSuite, 100, "test_borders" ) {}
 void runTest() { suite_StyleTestSuite.test_borders(); }
} testDescription_suite_StyleTestSuite_test_borders;

static class TestDescription_suite_StyleTestSuite_test_write_cell_xfs_1 : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_StyleTestSuite_test_write_cell_xfs_1() : CxxTest::RealTestDescription( Tests_StyleTestSuite, suiteDescription_StyleTestSuite, 109, "test_write_cell_xfs_1" ) {}
 void runTest() { suite_StyleTestSuite.test_write_cell_xfs_1(); }
} testDescription_suite_StyleTestSuite_test_write_cell_xfs_1;

static class TestDescription_suite_StyleTestSuite_test_alignment : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_StyleTestSuite_test_alignment() : CxxTest::RealTestDescription( Tests_StyleTestSuite, suiteDescription_StyleTestSuite, 123, "test_alignment" ) {}
 void runTest() { suite_StyleTestSuite.test_alignment(); }
} testDescription_suite_StyleTestSuite_test_alignment;

static class TestDescription_suite_StyleTestSuite_test_alignment_rotation : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_StyleTestSuite_test_alignment_rotation() : CxxTest::RealTestDescription( Tests_StyleTestSuite, suiteDescription_StyleTestSuite, 136, "test_alignment_rotation" ) {}
 void runTest() { suite_StyleTestSuite.test_alignment_rotation(); }
} testDescription_suite_StyleTestSuite_test_alignment_rotation;

static class TestDescription_suite_StyleTestSuite_test_format_comparisions : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_StyleTestSuite_test_format_comparisions() : CxxTest::RealTestDescription( Tests_StyleTestSuite, suiteDescription_StyleTestSuite, 151, "test_format_comparisions" ) {}
 void runTest() { suite_StyleTestSuite.test_format_comparisions(); }
} testDescription_suite_StyleTestSuite_test_format_comparisions;

static class TestDescription_suite_StyleTestSuite_test_builtin_format : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_StyleTestSuite_test_builtin_format() : CxxTest::RealTestDescription( Tests_StyleTestSuite, suiteDescription_StyleTestSuite, 166, "test_builtin_format" ) {}
 void runTest() { suite_StyleTestSuite.test_builtin_format(); }
} testDescription_suite_StyleTestSuite_test_builtin_format;

static class TestDescription_suite_StyleTestSuite_test_read_style : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_StyleTestSuite_test_read_style() : CxxTest::RealTestDescription( Tests_StyleTestSuite, suiteDescription_StyleTestSuite, 174, "test_read_style" ) {}
 void runTest() { suite_StyleTestSuite.test_read_style(); }
} testDescription_suite_StyleTestSuite_test_read_style;

static class TestDescription_suite_StyleTestSuite_test_read_cell_style : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_StyleTestSuite_test_read_cell_style() : CxxTest::RealTestDescription( Tests_StyleTestSuite, suiteDescription_StyleTestSuite, 191, "test_read_cell_style" ) {}
 void runTest() { suite_StyleTestSuite.test_read_cell_style(); }
} testDescription_suite_StyleTestSuite_test_read_cell_style;

#include "/Users/thomas/Development/xlnt/source/tests/ThemeTestSuite.h"

static ThemeTestSuite suite_ThemeTestSuite;

static CxxTest::List Tests_ThemeTestSuite = { 0, 0 };
CxxTest::StaticSuiteDescription suiteDescription_ThemeTestSuite( "../../source/tests/ThemeTestSuite.h", 8, "ThemeTestSuite", suite_ThemeTestSuite, Tests_ThemeTestSuite );

static class TestDescription_suite_ThemeTestSuite_test_write_theme : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ThemeTestSuite_test_write_theme() : CxxTest::RealTestDescription( Tests_ThemeTestSuite, suiteDescription_ThemeTestSuite, 16, "test_write_theme" ) {}
 void runTest() { suite_ThemeTestSuite.test_write_theme(); }
} testDescription_suite_ThemeTestSuite_test_write_theme;

#include "/Users/thomas/Development/xlnt/source/tests/UnicodeTestSuite.h"

static UnicodeTestSuite suite_UnicodeTestSuite;

static CxxTest::List Tests_UnicodeTestSuite = { 0, 0 };
CxxTest::StaticSuiteDescription suiteDescription_UnicodeTestSuite( "../../source/tests/UnicodeTestSuite.h", 8, "UnicodeTestSuite", suite_UnicodeTestSuite, Tests_UnicodeTestSuite );

static class TestDescription_suite_UnicodeTestSuite_test_read_workbook_with_unicode_character : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_UnicodeTestSuite_test_read_workbook_with_unicode_character() : CxxTest::RealTestDescription( Tests_UnicodeTestSuite, suiteDescription_UnicodeTestSuite, 16, "test_read_workbook_with_unicode_character" ) {}
 void runTest() { suite_UnicodeTestSuite.test_read_workbook_with_unicode_character(); }
} testDescription_suite_UnicodeTestSuite_test_read_workbook_with_unicode_character;

#include "/Users/thomas/Development/xlnt/source/tests/WorkbookTestSuite.h"

static WorkbookTestSuite suite_WorkbookTestSuite;

static CxxTest::List Tests_WorkbookTestSuite = { 0, 0 };
CxxTest::StaticSuiteDescription suiteDescription_WorkbookTestSuite( "../../source/tests/WorkbookTestSuite.h", 8, "WorkbookTestSuite", suite_WorkbookTestSuite, Tests_WorkbookTestSuite );

static class TestDescription_suite_WorkbookTestSuite_test_get_active_sheet : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorkbookTestSuite_test_get_active_sheet() : CxxTest::RealTestDescription( Tests_WorkbookTestSuite, suiteDescription_WorkbookTestSuite, 16, "test_get_active_sheet" ) {}
 void runTest() { suite_WorkbookTestSuite.test_get_active_sheet(); }
} testDescription_suite_WorkbookTestSuite_test_get_active_sheet;

static class TestDescription_suite_WorkbookTestSuite_test_create_sheet : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorkbookTestSuite_test_create_sheet() : CxxTest::RealTestDescription( Tests_WorkbookTestSuite, suiteDescription_WorkbookTestSuite, 23, "test_create_sheet" ) {}
 void runTest() { suite_WorkbookTestSuite.test_create_sheet(); }
} testDescription_suite_WorkbookTestSuite_test_create_sheet;

static class TestDescription_suite_WorkbookTestSuite_test_create_sheet_with_name : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorkbookTestSuite_test_create_sheet_with_name() : CxxTest::RealTestDescription( Tests_WorkbookTestSuite, suiteDescription_WorkbookTestSuite, 30, "test_create_sheet_with_name" ) {}
 void runTest() { suite_WorkbookTestSuite.test_create_sheet_with_name(); }
} testDescription_suite_WorkbookTestSuite_test_create_sheet_with_name;

static class TestDescription_suite_WorkbookTestSuite_test_add_correct_sheet : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorkbookTestSuite_test_add_correct_sheet() : CxxTest::RealTestDescription( Tests_WorkbookTestSuite, suiteDescription_WorkbookTestSuite, 37, "test_add_correct_sheet" ) {}
 void runTest() { suite_WorkbookTestSuite.test_add_correct_sheet(); }
} testDescription_suite_WorkbookTestSuite_test_add_correct_sheet;

static class TestDescription_suite_WorkbookTestSuite_test_create_sheet_readonly : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorkbookTestSuite_test_create_sheet_readonly() : CxxTest::RealTestDescription( Tests_WorkbookTestSuite, suiteDescription_WorkbookTestSuite, 45, "test_create_sheet_readonly" ) {}
 void runTest() { suite_WorkbookTestSuite.test_create_sheet_readonly(); }
} testDescription_suite_WorkbookTestSuite_test_create_sheet_readonly;

static class TestDescription_suite_WorkbookTestSuite_test_remove_sheet : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorkbookTestSuite_test_remove_sheet() : CxxTest::RealTestDescription( Tests_WorkbookTestSuite, suiteDescription_WorkbookTestSuite, 52, "test_remove_sheet" ) {}
 void runTest() { suite_WorkbookTestSuite.test_remove_sheet(); }
} testDescription_suite_WorkbookTestSuite_test_remove_sheet;

static class TestDescription_suite_WorkbookTestSuite_test_get_sheet_by_name : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorkbookTestSuite_test_get_sheet_by_name() : CxxTest::RealTestDescription( Tests_WorkbookTestSuite, suiteDescription_WorkbookTestSuite, 60, "test_get_sheet_by_name" ) {}
 void runTest() { suite_WorkbookTestSuite.test_get_sheet_by_name(); }
} testDescription_suite_WorkbookTestSuite_test_get_sheet_by_name;

static class TestDescription_suite_WorkbookTestSuite_test_get_index2 : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorkbookTestSuite_test_get_index2() : CxxTest::RealTestDescription( Tests_WorkbookTestSuite, suiteDescription_WorkbookTestSuite, 70, "test_get_index2" ) {}
 void runTest() { suite_WorkbookTestSuite.test_get_index2(); }
} testDescription_suite_WorkbookTestSuite_test_get_index2;

static class TestDescription_suite_WorkbookTestSuite_test_get_sheet_names : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorkbookTestSuite_test_get_sheet_names() : CxxTest::RealTestDescription( Tests_WorkbookTestSuite, suiteDescription_WorkbookTestSuite, 78, "test_get_sheet_names" ) {}
 void runTest() { suite_WorkbookTestSuite.test_get_sheet_names(); }
} testDescription_suite_WorkbookTestSuite_test_get_sheet_names;

static class TestDescription_suite_WorkbookTestSuite_test_get_named_ranges2 : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorkbookTestSuite_test_get_named_ranges2() : CxxTest::RealTestDescription( Tests_WorkbookTestSuite, suiteDescription_WorkbookTestSuite, 90, "test_get_named_ranges2" ) {}
 void runTest() { suite_WorkbookTestSuite.test_get_named_ranges2(); }
} testDescription_suite_WorkbookTestSuite_test_get_named_ranges2;

static class TestDescription_suite_WorkbookTestSuite_test_get_active_sheet2 : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorkbookTestSuite_test_get_active_sheet2() : CxxTest::RealTestDescription( Tests_WorkbookTestSuite, suiteDescription_WorkbookTestSuite, 95, "test_get_active_sheet2" ) {}
 void runTest() { suite_WorkbookTestSuite.test_get_active_sheet2(); }
} testDescription_suite_WorkbookTestSuite_test_get_active_sheet2;

static class TestDescription_suite_WorkbookTestSuite_test_create_sheet2 : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorkbookTestSuite_test_create_sheet2() : CxxTest::RealTestDescription( Tests_WorkbookTestSuite, suiteDescription_WorkbookTestSuite, 102, "test_create_sheet2" ) {}
 void runTest() { suite_WorkbookTestSuite.test_create_sheet2(); }
} testDescription_suite_WorkbookTestSuite_test_create_sheet2;

static class TestDescription_suite_WorkbookTestSuite_test_create_sheet_with_name2 : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorkbookTestSuite_test_create_sheet_with_name2() : CxxTest::RealTestDescription( Tests_WorkbookTestSuite, suiteDescription_WorkbookTestSuite, 109, "test_create_sheet_with_name2" ) {}
 void runTest() { suite_WorkbookTestSuite.test_create_sheet_with_name2(); }
} testDescription_suite_WorkbookTestSuite_test_create_sheet_with_name2;

static class TestDescription_suite_WorkbookTestSuite_test_add_correct_sheet2 : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorkbookTestSuite_test_add_correct_sheet2() : CxxTest::RealTestDescription( Tests_WorkbookTestSuite, suiteDescription_WorkbookTestSuite, 116, "test_add_correct_sheet2" ) {}
 void runTest() { suite_WorkbookTestSuite.test_add_correct_sheet2(); }
} testDescription_suite_WorkbookTestSuite_test_add_correct_sheet2;

static class TestDescription_suite_WorkbookTestSuite_test_add_incorrect_sheet2 : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorkbookTestSuite_test_add_incorrect_sheet2() : CxxTest::RealTestDescription( Tests_WorkbookTestSuite, suiteDescription_WorkbookTestSuite, 125, "test_add_incorrect_sheet2" ) {}
 void runTest() { suite_WorkbookTestSuite.test_add_incorrect_sheet2(); }
} testDescription_suite_WorkbookTestSuite_test_add_incorrect_sheet2;

static class TestDescription_suite_WorkbookTestSuite_test_create_sheet_readonly2 : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorkbookTestSuite_test_create_sheet_readonly2() : CxxTest::RealTestDescription( Tests_WorkbookTestSuite, suiteDescription_WorkbookTestSuite, 131, "test_create_sheet_readonly2" ) {}
 void runTest() { suite_WorkbookTestSuite.test_create_sheet_readonly2(); }
} testDescription_suite_WorkbookTestSuite_test_create_sheet_readonly2;

static class TestDescription_suite_WorkbookTestSuite_test_remove_sheet2 : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorkbookTestSuite_test_remove_sheet2() : CxxTest::RealTestDescription( Tests_WorkbookTestSuite, suiteDescription_WorkbookTestSuite, 138, "test_remove_sheet2" ) {}
 void runTest() { suite_WorkbookTestSuite.test_remove_sheet2(); }
} testDescription_suite_WorkbookTestSuite_test_remove_sheet2;

static class TestDescription_suite_WorkbookTestSuite_test_get_sheet_by_name2 : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorkbookTestSuite_test_get_sheet_by_name2() : CxxTest::RealTestDescription( Tests_WorkbookTestSuite, suiteDescription_WorkbookTestSuite, 146, "test_get_sheet_by_name2" ) {}
 void runTest() { suite_WorkbookTestSuite.test_get_sheet_by_name2(); }
} testDescription_suite_WorkbookTestSuite_test_get_sheet_by_name2;

static class TestDescription_suite_WorkbookTestSuite_test_get_index : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorkbookTestSuite_test_get_index() : CxxTest::RealTestDescription( Tests_WorkbookTestSuite, suiteDescription_WorkbookTestSuite, 156, "test_get_index" ) {}
 void runTest() { suite_WorkbookTestSuite.test_get_index(); }
} testDescription_suite_WorkbookTestSuite_test_get_index;

static class TestDescription_suite_WorkbookTestSuite_test_get_sheet_names2 : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorkbookTestSuite_test_get_sheet_names2() : CxxTest::RealTestDescription( Tests_WorkbookTestSuite, suiteDescription_WorkbookTestSuite, 164, "test_get_sheet_names2" ) {}
 void runTest() { suite_WorkbookTestSuite.test_get_sheet_names2(); }
} testDescription_suite_WorkbookTestSuite_test_get_sheet_names2;

static class TestDescription_suite_WorkbookTestSuite_test_get_named_ranges : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorkbookTestSuite_test_get_named_ranges() : CxxTest::RealTestDescription( Tests_WorkbookTestSuite, suiteDescription_WorkbookTestSuite, 176, "test_get_named_ranges" ) {}
 void runTest() { suite_WorkbookTestSuite.test_get_named_ranges(); }
} testDescription_suite_WorkbookTestSuite_test_get_named_ranges;

static class TestDescription_suite_WorkbookTestSuite_test_add_named_range : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorkbookTestSuite_test_add_named_range() : CxxTest::RealTestDescription( Tests_WorkbookTestSuite, suiteDescription_WorkbookTestSuite, 182, "test_add_named_range" ) {}
 void runTest() { suite_WorkbookTestSuite.test_add_named_range(); }
} testDescription_suite_WorkbookTestSuite_test_add_named_range;

static class TestDescription_suite_WorkbookTestSuite_test_get_named_range2 : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorkbookTestSuite_test_get_named_range2() : CxxTest::RealTestDescription( Tests_WorkbookTestSuite, suiteDescription_WorkbookTestSuite, 192, "test_get_named_range2" ) {}
 void runTest() { suite_WorkbookTestSuite.test_get_named_range2(); }
} testDescription_suite_WorkbookTestSuite_test_get_named_range2;

static class TestDescription_suite_WorkbookTestSuite_test_remove_named_range2 : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorkbookTestSuite_test_remove_named_range2() : CxxTest::RealTestDescription( Tests_WorkbookTestSuite, suiteDescription_WorkbookTestSuite, 202, "test_remove_named_range2" ) {}
 void runTest() { suite_WorkbookTestSuite.test_remove_named_range2(); }
} testDescription_suite_WorkbookTestSuite_test_remove_named_range2;

static class TestDescription_suite_WorkbookTestSuite_test_add_local_named_range2 : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorkbookTestSuite_test_add_local_named_range2() : CxxTest::RealTestDescription( Tests_WorkbookTestSuite, suiteDescription_WorkbookTestSuite, 213, "test_add_local_named_range2" ) {}
 void runTest() { suite_WorkbookTestSuite.test_add_local_named_range2(); }
} testDescription_suite_WorkbookTestSuite_test_add_local_named_range2;

static class TestDescription_suite_WorkbookTestSuite_test_write_regular_date : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorkbookTestSuite_test_write_regular_date() : CxxTest::RealTestDescription( Tests_WorkbookTestSuite, suiteDescription_WorkbookTestSuite, 226, "test_write_regular_date" ) {}
 void runTest() { suite_WorkbookTestSuite.test_write_regular_date(); }
} testDescription_suite_WorkbookTestSuite_test_write_regular_date;

static class TestDescription_suite_WorkbookTestSuite_test_write_regular_float : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorkbookTestSuite_test_write_regular_float() : CxxTest::RealTestDescription( Tests_WorkbookTestSuite, suiteDescription_WorkbookTestSuite, 244, "test_write_regular_float" ) {}
 void runTest() { suite_WorkbookTestSuite.test_write_regular_float(); }
} testDescription_suite_WorkbookTestSuite_test_write_regular_float;

static class TestDescription_suite_WorkbookTestSuite_test_bad_encoding2 : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorkbookTestSuite_test_bad_encoding2() : CxxTest::RealTestDescription( Tests_WorkbookTestSuite, suiteDescription_WorkbookTestSuite, 261, "test_bad_encoding2" ) {}
 void runTest() { suite_WorkbookTestSuite.test_bad_encoding2(); }
} testDescription_suite_WorkbookTestSuite_test_bad_encoding2;

static class TestDescription_suite_WorkbookTestSuite_test_good_encoding2 : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorkbookTestSuite_test_good_encoding2() : CxxTest::RealTestDescription( Tests_WorkbookTestSuite, suiteDescription_WorkbookTestSuite, 272, "test_good_encoding2" ) {}
 void runTest() { suite_WorkbookTestSuite.test_good_encoding2(); }
} testDescription_suite_WorkbookTestSuite_test_good_encoding2;

static class TestDescription_suite_WorkbookTestSuite_test_add_named_range2 : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorkbookTestSuite_test_add_named_range2() : CxxTest::RealTestDescription( Tests_WorkbookTestSuite, suiteDescription_WorkbookTestSuite, 282, "test_add_named_range2" ) {}
 void runTest() { suite_WorkbookTestSuite.test_add_named_range2(); }
} testDescription_suite_WorkbookTestSuite_test_add_named_range2;

static class TestDescription_suite_WorkbookTestSuite_test_get_named_range : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorkbookTestSuite_test_get_named_range() : CxxTest::RealTestDescription( Tests_WorkbookTestSuite, suiteDescription_WorkbookTestSuite, 292, "test_get_named_range" ) {}
 void runTest() { suite_WorkbookTestSuite.test_get_named_range(); }
} testDescription_suite_WorkbookTestSuite_test_get_named_range;

static class TestDescription_suite_WorkbookTestSuite_test_remove_named_range : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorkbookTestSuite_test_remove_named_range() : CxxTest::RealTestDescription( Tests_WorkbookTestSuite, suiteDescription_WorkbookTestSuite, 302, "test_remove_named_range" ) {}
 void runTest() { suite_WorkbookTestSuite.test_remove_named_range(); }
} testDescription_suite_WorkbookTestSuite_test_remove_named_range;

static class TestDescription_suite_WorkbookTestSuite_test_add_local_named_range : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorkbookTestSuite_test_add_local_named_range() : CxxTest::RealTestDescription( Tests_WorkbookTestSuite, suiteDescription_WorkbookTestSuite, 313, "test_add_local_named_range" ) {}
 void runTest() { suite_WorkbookTestSuite.test_add_local_named_range(); }
} testDescription_suite_WorkbookTestSuite_test_add_local_named_range;

static class TestDescription_suite_WorkbookTestSuite_test_write_regular_date2 : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorkbookTestSuite_test_write_regular_date2() : CxxTest::RealTestDescription( Tests_WorkbookTestSuite, suiteDescription_WorkbookTestSuite, 326, "test_write_regular_date2" ) {}
 void runTest() { suite_WorkbookTestSuite.test_write_regular_date2(); }
} testDescription_suite_WorkbookTestSuite_test_write_regular_date2;

static class TestDescription_suite_WorkbookTestSuite_test_write_regular_float2 : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorkbookTestSuite_test_write_regular_float2() : CxxTest::RealTestDescription( Tests_WorkbookTestSuite, suiteDescription_WorkbookTestSuite, 344, "test_write_regular_float2" ) {}
 void runTest() { suite_WorkbookTestSuite.test_write_regular_float2(); }
} testDescription_suite_WorkbookTestSuite_test_write_regular_float2;

static class TestDescription_suite_WorkbookTestSuite_test_bad_encoding : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorkbookTestSuite_test_bad_encoding() : CxxTest::RealTestDescription( Tests_WorkbookTestSuite, suiteDescription_WorkbookTestSuite, 363, "test_bad_encoding" ) {}
 void runTest() { suite_WorkbookTestSuite.test_bad_encoding(); }
} testDescription_suite_WorkbookTestSuite_test_bad_encoding;

static class TestDescription_suite_WorkbookTestSuite_test_good_encoding : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorkbookTestSuite_test_good_encoding() : CxxTest::RealTestDescription( Tests_WorkbookTestSuite, suiteDescription_WorkbookTestSuite, 373, "test_good_encoding" ) {}
 void runTest() { suite_WorkbookTestSuite.test_good_encoding(); }
} testDescription_suite_WorkbookTestSuite_test_good_encoding;

#include "/Users/thomas/Development/xlnt/source/tests/WorksheetTestSuite.h"

static WorksheetTestSuite suite_WorksheetTestSuite;

static CxxTest::List Tests_WorksheetTestSuite = { 0, 0 };
CxxTest::StaticSuiteDescription suiteDescription_WorksheetTestSuite( "../../source/tests/WorksheetTestSuite.h", 9, "WorksheetTestSuite", suite_WorksheetTestSuite, Tests_WorksheetTestSuite );

static class TestDescription_suite_WorksheetTestSuite_test_new_worksheet : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorksheetTestSuite_test_new_worksheet() : CxxTest::RealTestDescription( Tests_WorksheetTestSuite, suiteDescription_WorksheetTestSuite, 16, "test_new_worksheet" ) {}
 void runTest() { suite_WorksheetTestSuite.test_new_worksheet(); }
} testDescription_suite_WorksheetTestSuite_test_new_worksheet;

static class TestDescription_suite_WorksheetTestSuite_test_new_sheet_name : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorksheetTestSuite_test_new_sheet_name() : CxxTest::RealTestDescription( Tests_WorksheetTestSuite, suiteDescription_WorksheetTestSuite, 22, "test_new_sheet_name" ) {}
 void runTest() { suite_WorksheetTestSuite.test_new_sheet_name(); }
} testDescription_suite_WorksheetTestSuite_test_new_sheet_name;

static class TestDescription_suite_WorksheetTestSuite_test_get_cell : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorksheetTestSuite_test_get_cell() : CxxTest::RealTestDescription( Tests_WorksheetTestSuite, suiteDescription_WorksheetTestSuite, 30, "test_get_cell" ) {}
 void runTest() { suite_WorksheetTestSuite.test_get_cell(); }
} testDescription_suite_WorksheetTestSuite_test_get_cell;

static class TestDescription_suite_WorksheetTestSuite_test_set_bad_title : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorksheetTestSuite_test_set_bad_title() : CxxTest::RealTestDescription( Tests_WorksheetTestSuite, suiteDescription_WorksheetTestSuite, 37, "test_set_bad_title" ) {}
 void runTest() { suite_WorksheetTestSuite.test_set_bad_title(); }
} testDescription_suite_WorksheetTestSuite_test_set_bad_title;

static class TestDescription_suite_WorksheetTestSuite_test_set_bad_title_character : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorksheetTestSuite_test_set_bad_title_character() : CxxTest::RealTestDescription( Tests_WorksheetTestSuite, suiteDescription_WorksheetTestSuite, 43, "test_set_bad_title_character" ) {}
 void runTest() { suite_WorksheetTestSuite.test_set_bad_title_character(); }
} testDescription_suite_WorksheetTestSuite_test_set_bad_title_character;

static class TestDescription_suite_WorksheetTestSuite_test_worksheet_dimension : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorksheetTestSuite_test_worksheet_dimension() : CxxTest::RealTestDescription( Tests_WorksheetTestSuite, suiteDescription_WorksheetTestSuite, 54, "test_worksheet_dimension" ) {}
 void runTest() { suite_WorksheetTestSuite.test_worksheet_dimension(); }
} testDescription_suite_WorksheetTestSuite_test_worksheet_dimension;

static class TestDescription_suite_WorksheetTestSuite_test_worksheet_range : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorksheetTestSuite_test_worksheet_range() : CxxTest::RealTestDescription( Tests_WorksheetTestSuite, suiteDescription_WorksheetTestSuite, 62, "test_worksheet_range" ) {}
 void runTest() { suite_WorksheetTestSuite.test_worksheet_range(); }
} testDescription_suite_WorksheetTestSuite_test_worksheet_range;

static class TestDescription_suite_WorksheetTestSuite_test_worksheet_named_range : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorksheetTestSuite_test_worksheet_named_range() : CxxTest::RealTestDescription( Tests_WorksheetTestSuite, suiteDescription_WorksheetTestSuite, 70, "test_worksheet_named_range" ) {}
 void runTest() { suite_WorksheetTestSuite.test_worksheet_named_range(); }
} testDescription_suite_WorksheetTestSuite_test_worksheet_named_range;

static class TestDescription_suite_WorksheetTestSuite_test_bad_named_range : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorksheetTestSuite_test_bad_named_range() : CxxTest::RealTestDescription( Tests_WorksheetTestSuite, suiteDescription_WorksheetTestSuite, 80, "test_bad_named_range" ) {}
 void runTest() { suite_WorksheetTestSuite.test_bad_named_range(); }
} testDescription_suite_WorksheetTestSuite_test_bad_named_range;

static class TestDescription_suite_WorksheetTestSuite_test_named_range_wrong_sheet : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorksheetTestSuite_test_named_range_wrong_sheet() : CxxTest::RealTestDescription( Tests_WorksheetTestSuite, suiteDescription_WorksheetTestSuite, 86, "test_named_range_wrong_sheet" ) {}
 void runTest() { suite_WorksheetTestSuite.test_named_range_wrong_sheet(); }
} testDescription_suite_WorksheetTestSuite_test_named_range_wrong_sheet;

static class TestDescription_suite_WorksheetTestSuite_test_cell_offset : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorksheetTestSuite_test_cell_offset() : CxxTest::RealTestDescription( Tests_WorksheetTestSuite, suiteDescription_WorksheetTestSuite, 94, "test_cell_offset" ) {}
 void runTest() { suite_WorksheetTestSuite.test_cell_offset(); }
} testDescription_suite_WorksheetTestSuite_test_cell_offset;

static class TestDescription_suite_WorksheetTestSuite_test_range_offset : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorksheetTestSuite_test_range_offset() : CxxTest::RealTestDescription( Tests_WorksheetTestSuite, suiteDescription_WorksheetTestSuite, 100, "test_range_offset" ) {}
 void runTest() { suite_WorksheetTestSuite.test_range_offset(); }
} testDescription_suite_WorksheetTestSuite_test_range_offset;

static class TestDescription_suite_WorksheetTestSuite_test_cell_alternate_coordinates : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorksheetTestSuite_test_cell_alternate_coordinates() : CxxTest::RealTestDescription( Tests_WorksheetTestSuite, suiteDescription_WorksheetTestSuite, 109, "test_cell_alternate_coordinates" ) {}
 void runTest() { suite_WorksheetTestSuite.test_cell_alternate_coordinates(); }
} testDescription_suite_WorksheetTestSuite_test_cell_alternate_coordinates;

static class TestDescription_suite_WorksheetTestSuite_test_cell_range_name : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorksheetTestSuite_test_cell_range_name() : CxxTest::RealTestDescription( Tests_WorksheetTestSuite, suiteDescription_WorksheetTestSuite, 116, "test_cell_range_name" ) {}
 void runTest() { suite_WorksheetTestSuite.test_cell_range_name(); }
} testDescription_suite_WorksheetTestSuite_test_cell_range_name;

static class TestDescription_suite_WorksheetTestSuite_test_garbage_collect : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorksheetTestSuite_test_garbage_collect() : CxxTest::RealTestDescription( Tests_WorksheetTestSuite, suiteDescription_WorksheetTestSuite, 128, "test_garbage_collect" ) {}
 void runTest() { suite_WorksheetTestSuite.test_garbage_collect(); }
} testDescription_suite_WorksheetTestSuite_test_garbage_collect;

static class TestDescription_suite_WorksheetTestSuite_test_hyperlink_relationships : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorksheetTestSuite_test_hyperlink_relationships() : CxxTest::RealTestDescription( Tests_WorksheetTestSuite, suiteDescription_WorksheetTestSuite, 150, "test_hyperlink_relationships" ) {}
 void runTest() { suite_WorksheetTestSuite.test_hyperlink_relationships(); }
} testDescription_suite_WorksheetTestSuite_test_hyperlink_relationships;

static class TestDescription_suite_WorksheetTestSuite_test_bad_relationship_type : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorksheetTestSuite_test_bad_relationship_type() : CxxTest::RealTestDescription( Tests_WorksheetTestSuite, suiteDescription_WorksheetTestSuite, 171, "test_bad_relationship_type" ) {}
 void runTest() { suite_WorksheetTestSuite.test_bad_relationship_type(); }
} testDescription_suite_WorksheetTestSuite_test_bad_relationship_type;

static class TestDescription_suite_WorksheetTestSuite_test_append_list : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorksheetTestSuite_test_append_list() : CxxTest::RealTestDescription( Tests_WorksheetTestSuite, suiteDescription_WorksheetTestSuite, 176, "test_append_list" ) {}
 void runTest() { suite_WorksheetTestSuite.test_append_list(); }
} testDescription_suite_WorksheetTestSuite_test_append_list;

static class TestDescription_suite_WorksheetTestSuite_test_append_dict_letter : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorksheetTestSuite_test_append_dict_letter() : CxxTest::RealTestDescription( Tests_WorksheetTestSuite, suiteDescription_WorksheetTestSuite, 186, "test_append_dict_letter" ) {}
 void runTest() { suite_WorksheetTestSuite.test_append_dict_letter(); }
} testDescription_suite_WorksheetTestSuite_test_append_dict_letter;

static class TestDescription_suite_WorksheetTestSuite_test_append_dict_index : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorksheetTestSuite_test_append_dict_index() : CxxTest::RealTestDescription( Tests_WorksheetTestSuite, suiteDescription_WorksheetTestSuite, 196, "test_append_dict_index" ) {}
 void runTest() { suite_WorksheetTestSuite.test_append_dict_index(); }
} testDescription_suite_WorksheetTestSuite_test_append_dict_index;

static class TestDescription_suite_WorksheetTestSuite_test_append_2d_list : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorksheetTestSuite_test_append_2d_list() : CxxTest::RealTestDescription( Tests_WorksheetTestSuite, suiteDescription_WorksheetTestSuite, 206, "test_append_2d_list" ) {}
 void runTest() { suite_WorksheetTestSuite.test_append_2d_list(); }
} testDescription_suite_WorksheetTestSuite_test_append_2d_list;

static class TestDescription_suite_WorksheetTestSuite_test_rows : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorksheetTestSuite_test_rows() : CxxTest::RealTestDescription( Tests_WorksheetTestSuite, suiteDescription_WorksheetTestSuite, 221, "test_rows" ) {}
 void runTest() { suite_WorksheetTestSuite.test_rows(); }
} testDescription_suite_WorksheetTestSuite_test_rows;

static class TestDescription_suite_WorksheetTestSuite_test_cols : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorksheetTestSuite_test_cols() : CxxTest::RealTestDescription( Tests_WorksheetTestSuite, suiteDescription_WorksheetTestSuite, 236, "test_cols" ) {}
 void runTest() { suite_WorksheetTestSuite.test_cols(); }
} testDescription_suite_WorksheetTestSuite_test_cols;

static class TestDescription_suite_WorksheetTestSuite_test_auto_filter : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorksheetTestSuite_test_auto_filter() : CxxTest::RealTestDescription( Tests_WorksheetTestSuite, suiteDescription_WorksheetTestSuite, 251, "test_auto_filter" ) {}
 void runTest() { suite_WorksheetTestSuite.test_auto_filter(); }
} testDescription_suite_WorksheetTestSuite_test_auto_filter;

static class TestDescription_suite_WorksheetTestSuite_test_page_margins : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorksheetTestSuite_test_page_margins() : CxxTest::RealTestDescription( Tests_WorksheetTestSuite, suiteDescription_WorksheetTestSuite, 265, "test_page_margins" ) {}
 void runTest() { suite_WorksheetTestSuite.test_page_margins(); }
} testDescription_suite_WorksheetTestSuite_test_page_margins;

static class TestDescription_suite_WorksheetTestSuite_test_merge : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorksheetTestSuite_test_merge() : CxxTest::RealTestDescription( Tests_WorksheetTestSuite, suiteDescription_WorksheetTestSuite, 285, "test_merge" ) {}
 void runTest() { suite_WorksheetTestSuite.test_merge(); }
} testDescription_suite_WorksheetTestSuite_test_merge;

static class TestDescription_suite_WorksheetTestSuite_test_freeze : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorksheetTestSuite_test_freeze() : CxxTest::RealTestDescription( Tests_WorksheetTestSuite, suiteDescription_WorksheetTestSuite, 306, "test_freeze" ) {}
 void runTest() { suite_WorksheetTestSuite.test_freeze(); }
} testDescription_suite_WorksheetTestSuite_test_freeze;

static class TestDescription_suite_WorksheetTestSuite_test_printer_settings : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WorksheetTestSuite_test_printer_settings() : CxxTest::RealTestDescription( Tests_WorksheetTestSuite, suiteDescription_WorksheetTestSuite, 323, "test_printer_settings" ) {}
 void runTest() { suite_WorksheetTestSuite.test_printer_settings(); }
} testDescription_suite_WorksheetTestSuite_test_printer_settings;

#include "/Users/thomas/Development/xlnt/source/tests/WriteTestSuite.h"

static WriteTestSuite suite_WriteTestSuite;

static CxxTest::List Tests_WriteTestSuite = { 0, 0 };
CxxTest::StaticSuiteDescription suiteDescription_WriteTestSuite( "../../source/tests/WriteTestSuite.h", 9, "WriteTestSuite", suite_WriteTestSuite, Tests_WriteTestSuite );

static class TestDescription_suite_WriteTestSuite_test_write_empty_workbook : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WriteTestSuite_test_write_empty_workbook() : CxxTest::RealTestDescription( Tests_WriteTestSuite, suiteDescription_WriteTestSuite, 19, "test_write_empty_workbook" ) {}
 void runTest() { suite_WriteTestSuite.test_write_empty_workbook(); }
} testDescription_suite_WriteTestSuite_test_write_empty_workbook;

static class TestDescription_suite_WriteTestSuite_test_write_virtual_workbook : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WriteTestSuite_test_write_virtual_workbook() : CxxTest::RealTestDescription( Tests_WriteTestSuite, suiteDescription_WriteTestSuite, 28, "test_write_virtual_workbook" ) {}
 void runTest() { suite_WriteTestSuite.test_write_virtual_workbook(); }
} testDescription_suite_WriteTestSuite_test_write_virtual_workbook;

static class TestDescription_suite_WriteTestSuite_test_write_workbook_rels : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WriteTestSuite_test_write_workbook_rels() : CxxTest::RealTestDescription( Tests_WriteTestSuite, suiteDescription_WriteTestSuite, 36, "test_write_workbook_rels" ) {}
 void runTest() { suite_WriteTestSuite.test_write_workbook_rels(); }
} testDescription_suite_WriteTestSuite_test_write_workbook_rels;

static class TestDescription_suite_WriteTestSuite_test_write_workbook : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WriteTestSuite_test_write_workbook() : CxxTest::RealTestDescription( Tests_WriteTestSuite, suiteDescription_WriteTestSuite, 44, "test_write_workbook" ) {}
 void runTest() { suite_WriteTestSuite.test_write_workbook(); }
} testDescription_suite_WriteTestSuite_test_write_workbook;

static class TestDescription_suite_WriteTestSuite_test_write_string_table : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WriteTestSuite_test_write_string_table() : CxxTest::RealTestDescription( Tests_WriteTestSuite, suiteDescription_WriteTestSuite, 53, "test_write_string_table" ) {}
 void runTest() { suite_WriteTestSuite.test_write_string_table(); }
} testDescription_suite_WriteTestSuite_test_write_string_table;

static class TestDescription_suite_WriteTestSuite_test_write_worksheet : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WriteTestSuite_test_write_worksheet() : CxxTest::RealTestDescription( Tests_WriteTestSuite, suiteDescription_WriteTestSuite, 61, "test_write_worksheet" ) {}
 void runTest() { suite_WriteTestSuite.test_write_worksheet(); }
} testDescription_suite_WriteTestSuite_test_write_worksheet;

static class TestDescription_suite_WriteTestSuite_test_write_hidden_worksheet : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WriteTestSuite_test_write_hidden_worksheet() : CxxTest::RealTestDescription( Tests_WriteTestSuite, suiteDescription_WriteTestSuite, 71, "test_write_hidden_worksheet" ) {}
 void runTest() { suite_WriteTestSuite.test_write_hidden_worksheet(); }
} testDescription_suite_WriteTestSuite_test_write_hidden_worksheet;

static class TestDescription_suite_WriteTestSuite_test_write_bool : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WriteTestSuite_test_write_bool() : CxxTest::RealTestDescription( Tests_WriteTestSuite, suiteDescription_WriteTestSuite, 82, "test_write_bool" ) {}
 void runTest() { suite_WriteTestSuite.test_write_bool(); }
} testDescription_suite_WriteTestSuite_test_write_bool;

static class TestDescription_suite_WriteTestSuite_test_write_formula : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WriteTestSuite_test_write_formula() : CxxTest::RealTestDescription( Tests_WriteTestSuite, suiteDescription_WriteTestSuite, 93, "test_write_formula" ) {}
 void runTest() { suite_WriteTestSuite.test_write_formula(); }
} testDescription_suite_WriteTestSuite_test_write_formula;

static class TestDescription_suite_WriteTestSuite_test_write_style : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WriteTestSuite_test_write_style() : CxxTest::RealTestDescription( Tests_WriteTestSuite, suiteDescription_WriteTestSuite, 105, "test_write_style" ) {}
 void runTest() { suite_WriteTestSuite.test_write_style(); }
} testDescription_suite_WriteTestSuite_test_write_style;

static class TestDescription_suite_WriteTestSuite_test_write_height : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WriteTestSuite_test_write_height() : CxxTest::RealTestDescription( Tests_WriteTestSuite, suiteDescription_WriteTestSuite, 116, "test_write_height" ) {}
 void runTest() { suite_WriteTestSuite.test_write_height(); }
} testDescription_suite_WriteTestSuite_test_write_height;

static class TestDescription_suite_WriteTestSuite_test_write_hyperlink : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WriteTestSuite_test_write_hyperlink() : CxxTest::RealTestDescription( Tests_WriteTestSuite, suiteDescription_WriteTestSuite, 127, "test_write_hyperlink" ) {}
 void runTest() { suite_WriteTestSuite.test_write_hyperlink(); }
} testDescription_suite_WriteTestSuite_test_write_hyperlink;

static class TestDescription_suite_WriteTestSuite_test_write_hyperlink_rels : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WriteTestSuite_test_write_hyperlink_rels() : CxxTest::RealTestDescription( Tests_WriteTestSuite, suiteDescription_WriteTestSuite, 138, "test_write_hyperlink_rels" ) {}
 void runTest() { suite_WriteTestSuite.test_write_hyperlink_rels(); }
} testDescription_suite_WriteTestSuite_test_write_hyperlink_rels;

static class TestDescription_suite_WriteTestSuite_test_hyperlink_value : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WriteTestSuite_test_hyperlink_value() : CxxTest::RealTestDescription( Tests_WriteTestSuite, suiteDescription_WriteTestSuite, 154, "test_hyperlink_value" ) {}
 void runTest() { suite_WriteTestSuite.test_hyperlink_value(); }
} testDescription_suite_WriteTestSuite_test_hyperlink_value;

static class TestDescription_suite_WriteTestSuite_test_write_auto_filter : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WriteTestSuite_test_write_auto_filter() : CxxTest::RealTestDescription( Tests_WriteTestSuite, suiteDescription_WriteTestSuite, 164, "test_write_auto_filter" ) {}
 void runTest() { suite_WriteTestSuite.test_write_auto_filter(); }
} testDescription_suite_WriteTestSuite_test_write_auto_filter;

static class TestDescription_suite_WriteTestSuite_test_freeze_panes_horiz : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WriteTestSuite_test_freeze_panes_horiz() : CxxTest::RealTestDescription( Tests_WriteTestSuite, suiteDescription_WriteTestSuite, 179, "test_freeze_panes_horiz" ) {}
 void runTest() { suite_WriteTestSuite.test_freeze_panes_horiz(); }
} testDescription_suite_WriteTestSuite_test_freeze_panes_horiz;

static class TestDescription_suite_WriteTestSuite_test_freeze_panes_vert : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WriteTestSuite_test_freeze_panes_vert() : CxxTest::RealTestDescription( Tests_WriteTestSuite, suiteDescription_WriteTestSuite, 190, "test_freeze_panes_vert" ) {}
 void runTest() { suite_WriteTestSuite.test_freeze_panes_vert(); }
} testDescription_suite_WriteTestSuite_test_freeze_panes_vert;

static class TestDescription_suite_WriteTestSuite_test_freeze_panes_both : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WriteTestSuite_test_freeze_panes_both() : CxxTest::RealTestDescription( Tests_WriteTestSuite, suiteDescription_WriteTestSuite, 201, "test_freeze_panes_both" ) {}
 void runTest() { suite_WriteTestSuite.test_freeze_panes_both(); }
} testDescription_suite_WriteTestSuite_test_freeze_panes_both;

static class TestDescription_suite_WriteTestSuite_test_long_number : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WriteTestSuite_test_long_number() : CxxTest::RealTestDescription( Tests_WriteTestSuite, suiteDescription_WriteTestSuite, 212, "test_long_number" ) {}
 void runTest() { suite_WriteTestSuite.test_long_number(); }
} testDescription_suite_WriteTestSuite_test_long_number;

static class TestDescription_suite_WriteTestSuite_test_decimal : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WriteTestSuite_test_decimal() : CxxTest::RealTestDescription( Tests_WriteTestSuite, suiteDescription_WriteTestSuite, 222, "test_decimal" ) {}
 void runTest() { suite_WriteTestSuite.test_decimal(); }
} testDescription_suite_WriteTestSuite_test_decimal;

static class TestDescription_suite_WriteTestSuite_test_short_number : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_WriteTestSuite_test_short_number() : CxxTest::RealTestDescription( Tests_WriteTestSuite, suiteDescription_WriteTestSuite, 232, "test_short_number" ) {}
 void runTest() { suite_WriteTestSuite.test_short_number(); }
} testDescription_suite_WriteTestSuite_test_short_number;

#include "/Users/thomas/Development/xlnt/source/tests/ZipFileTestSuite.h"

static ZipFileTestSuite suite_ZipFileTestSuite;

static CxxTest::List Tests_ZipFileTestSuite = { 0, 0 };
CxxTest::StaticSuiteDescription suiteDescription_ZipFileTestSuite( "../../source/tests/ZipFileTestSuite.h", 8, "ZipFileTestSuite", suite_ZipFileTestSuite, Tests_ZipFileTestSuite );

static class TestDescription_suite_ZipFileTestSuite_test_existing_package : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ZipFileTestSuite_test_existing_package() : CxxTest::RealTestDescription( Tests_ZipFileTestSuite, suiteDescription_ZipFileTestSuite, 22, "test_existing_package" ) {}
 void runTest() { suite_ZipFileTestSuite.test_existing_package(); }
} testDescription_suite_ZipFileTestSuite_test_existing_package;

static class TestDescription_suite_ZipFileTestSuite_test_new_package : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ZipFileTestSuite_test_new_package() : CxxTest::RealTestDescription( Tests_ZipFileTestSuite, suiteDescription_ZipFileTestSuite, 28, "test_new_package" ) {}
 void runTest() { suite_ZipFileTestSuite.test_new_package(); }
} testDescription_suite_ZipFileTestSuite_test_new_package;

static class TestDescription_suite_ZipFileTestSuite_test_read_text : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ZipFileTestSuite_test_read_text() : CxxTest::RealTestDescription( Tests_ZipFileTestSuite, suiteDescription_ZipFileTestSuite, 38, "test_read_text" ) {}
 void runTest() { suite_ZipFileTestSuite.test_read_text(); }
} testDescription_suite_ZipFileTestSuite_test_read_text;

static class TestDescription_suite_ZipFileTestSuite_test_write_text : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ZipFileTestSuite_test_write_text() : CxxTest::RealTestDescription( Tests_ZipFileTestSuite, suiteDescription_ZipFileTestSuite, 45, "test_write_text" ) {}
 void runTest() { suite_ZipFileTestSuite.test_write_text(); }
} testDescription_suite_ZipFileTestSuite_test_write_text;

static class TestDescription_suite_ZipFileTestSuite_test_read_xml : public CxxTest::RealTestDescription {
public:
 TestDescription_suite_ZipFileTestSuite_test_read_xml() : CxxTest::RealTestDescription( Tests_ZipFileTestSuite, suiteDescription_ZipFileTestSuite, 59, "test_read_xml" ) {}
 void runTest() { suite_ZipFileTestSuite.test_read_xml(); }
} testDescription_suite_ZipFileTestSuite_test_read_xml;

#include <cxxtest/Root.cpp>
const char* CxxTest::RealWorldDescription::_worldName = "cxxtest";
