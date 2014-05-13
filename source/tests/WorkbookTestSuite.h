#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include "../xlnt.h"

class WorkbookTestSuite : public CxxTest::TestSuite
{
public:
    WorkbookTestSuite()
    {

    }

    void test_get_active_sheet()
    {
        /*xlnt::workbook wb;
        auto active_sheet = wb.get_active_sheet();
        TS_ASSERT_EQUALS(active_sheet, wb.worksheets[0]);*/
    }

    void test_create_sheet()
    {
        /*xlnt::workbook wb;
        new_sheet = wb.create_sheet(0);
        TS_ASSERT_EQUALS(new_sheet, wb.worksheets[0]);*/
    }

    void test_create_sheet_with_name()
    {
        /*xlnt::workbook wb;
        new_sheet = wb.create_sheet(0, title = "LikeThisName");
        TS_ASSERT_EQUALS(new_sheet, wb.worksheets[0]);*/
    }

    void test_add_correct_sheet()
    {
        /*xlnt::workbook wb;
        new_sheet = wb.create_sheet(0);
        wb.add_sheet(new_sheet);
        TS_ASSERT_EQUALS(new_sheet, wb.worksheets[2]);*/
    }

    void test_create_sheet_readonly()
    {
        /*xlnt::workbook wb;
        wb._set_optimized_read();
        wb.create_sheet();*/
    }

    void test_remove_sheet()
    {
        /*xlnt::workbook wb;
        new_sheet = wb.create_sheet(0);
        wb.remove_sheet(new_sheet);
        assert new_sheet not in wb.worksheets;*/
    }

    void test_get_sheet_by_name()
    {
        /*xlnt::workbook wb;
        new_sheet = wb.create_sheet();
        title = "my sheet";
        new_sheet.title = title;
        found_sheet = wb.get_sheet_by_name(title);
        TS_ASSERT_EQUALS(new_sheet, found_sheet);*/
    }

    void test_get_index2()
    {
        /*xlnt::workbook wb;
        new_sheet = wb.create_sheet(0);
        sheet_index = wb.get_index(new_sheet);
        TS_ASSERT_EQUALS(sheet_index, 0);*/
    }

    void test_get_sheet_names()
    {
        /*xlnt::workbook wb;
        names = ["Sheet", "Sheet1", "Sheet2", "Sheet3", "Sheet4", "Sheet5"];
        for(auto count : range(5))
        {
            wb.create_sheet(0)
                actual_names = wb.get_sheet_names()
                TS_ASSERT_EQUALS(sorted(actual_names), sorted(names))
        }*/
    }

    void test_get_named_ranges2()
    {
        /*xlnt::workbook wb;
        TS_ASSERT_EQUALS(wb.get_named_ranges(), wb._named_ranges);*/
    }
    void test_get_active_sheet2()
    {
        /*xlnt::workbook wb;
        active_sheet = wb.get_active_sheet();
        TS_ASSERT_EQUALS(active_sheet, wb.worksheets[0]);*/
    }

    void test_create_sheet2()
    {
        /*xlnt::workbook wb;
        new_sheet = wb.create_sheet(0);
        TS_ASSERT_EQUALS(new_sheet, wb.worksheets[0]);*/
    }

    void test_create_sheet_with_name2()
    {
        /*xlnt::workbook wb;
        new_sheet = wb.create_sheet(0, title = "LikeThisName");
        TS_ASSERT_EQUALS(new_sheet, wb.worksheets[0]);*/
    }

    void test_add_correct_sheet2()
    {
        //xlnt::workbook wb;
        //new_sheet = wb.create_sheet(0);
        //wb.add_sheet(new_sheet);
        //TS_ASSERT_EQUALS(new_sheet, wb.worksheets[2]);
    }

        //@raises(AssertionError)
    void test_add_incorrect_sheet2()
    {
        //xlnt::workbook wb;
        //TS_ASSERT_THROWS(AssertionError, wb.add_sheet("Test"))
    }

    void test_create_sheet_readonly2()
    {
        //xlnt::workbook wb;
        //wb._set_optimized_read();
        //TS_ASSERT_THROWS(wb.create_sheet(), ReadOnlyWorkbook);
    }

    void test_remove_sheet2()
    {
        //xlnt::workbook wb;
        //new_sheet = wb.create_sheet(0);
        //wb.remove_sheet(new_sheet);
        //assert new_sheet not in wb.worksheets;
    }

    void test_get_sheet_by_name2()
    {
        //xlnt::workbook wb;
        //new_sheet = wb.create_sheet();
        //title = "my sheet";
        //new_sheet.title = title;
        //found_sheet = wb.get_sheet_by_name(title);
        //TS_ASSERT_EQUALS(new_sheet, found_sheet);
    }

    void test_get_index()
    {
        //xlnt::workbook wb;
        //new_sheet = wb.create_sheet(0);
        //sheet_index = wb.get_index(new_sheet);
        //TS_ASSERT_EQUALS(sheet_index, 0);
    }

    void test_get_sheet_names2()
    {
        //xlnt::workbook wb;
        //names = ["Sheet", "Sheet1", "Sheet2", "Sheet3", "Sheet4", "Sheet5"];
        //for(auto count in range(5))
        //{
        //    wb.create_sheet(0)
        //        actual_names = wb.get_sheet_names()
        //        TS_ASSERT_EQUALS(sorted(actual_names), sorted(names))
        //}
    }

    void test_get_named_ranges()
    {
        //xlnt::workbook wb;
        //TS_ASSERT_EQUALS(wb.get_named_ranges(), wb._named_ranges);
    }

    void test_add_named_range()
    {
        //xlnt::workbook wb;
        //new_sheet = wb.create_sheet();
        //named_range = NamedRange("test_nr", [(new_sheet, "A1")]);
        //wb.add_named_range(named_range);
        //named_ranges_list = wb.get_named_ranges();
        //assert named_range in named_ranges_list;
    }

    void test_get_named_range2()
    {
        //xlnt::workbook wb;
        //new_sheet = wb.create_sheet();
        //named_range = NamedRange("test_nr", [(new_sheet, "A1")]);
        //wb.add_named_range(named_range);
        //found_named_range = wb.get_named_range("test_nr");
        //TS_ASSERT_EQUALS(named_range, found_named_range);
    }

    void test_remove_named_range2()
    {
        //xlnt::workbook wb;
        //new_sheet = wb.create_sheet();
        //named_range = NamedRange("test_nr", [(new_sheet, "A1")]);
        //wb.add_named_range(named_range);
        //wb.remove_named_range(named_range);
        //named_ranges_list = wb.get_named_ranges();
        //assert named_range not in named_ranges_list;
    }

    void test_add_local_named_range2()
    {
        //make_tmpdir();
        //xlnt::workbook wb;
        //new_sheet = wb.create_sheet();
        //named_range = NamedRange("test_nr", [(new_sheet, "A1")]);
        //named_range.scope = new_sheet;
        //wb.add_named_range(named_range);
        //dest_filename = osp.join(TMPDIR, "local_named_range_book.xlsx");
        //wb.save(dest_filename);
        //clean_tmpdir();
    }

    void test_write_regular_date()
    {
        //make_tmpdir();
        //today = datetime.datetime(2010, 1, 18, 14, 15, 20, 1600);

        //book = Workbook();
        //sheet = book.get_active_sheet();
        //sheet.cell("A1").value = today;
        //dest_filename = osp.join(TMPDIR, "date_read_write_issue.xlsx");
        //book.save(dest_filename);

        //test_book = load_workbook(dest_filename);
        //test_sheet = test_book.get_active_sheet();

        //TS_ASSERT_EQUALS(test_sheet.cell("A1"), today);
        //clean_tmpdir();
    }

    void test_write_regular_float()
    {
        //make_tmpdir();
        //float float_value = 1.0 / 3.0;
        //book = Workbook();
        //sheet = book.get_active_sheet();
        //sheet.cell("A1").value = float_value;
        //dest_filename = osp.join(TMPDIR, "float_read_write_issue.xlsx");
        //book.save(dest_filename);

        //test_book = load_workbook(dest_filename);
        //test_sheet = test_book.get_active_sheet();

        //TS_ASSERT_EQUALS(test_sheet.cell("A1").value, float_value);*/
        //clean_tmpdir();
    }

    void test_bad_encoding2()
    {
        //char pound = 163;
        //std::string test_string = ("Compound Value (" + std::string(1, pound) + ")").encode("latin1");

        //xlnt::workbook utf_book;
        //xlnt::worksheet utf_sheet = utf_book.get_active_sheet();

        //TS_ASSERT_THROWS(UnicodeDecode, utf_sheet.cell("A1") = test_string);
    }

    void test_good_encoding2()
    {
        //char pound = 163;
        //std::string test_string = ("Compound Value (" + std::string(1, pound) + ")").encode("latin1");

        //lat_book = Workbook(encoding = "latin1");
        //lat_sheet = lat_book.get_active_sheet();
        //lat_sheet.cell("A1").value = test_string;
    }

    void test_add_named_range2()
    {
        //wb = Workbook();
        //new_sheet = wb.create_sheet();
        //named_range = NamedRange("test_nr", [(new_sheet, "A1")]);
        //wb.add_named_range(named_range);
        //named_ranges_list = wb.get_named_ranges();
        //assert named_range in named_ranges_list;
    }

    void test_get_named_range()
    {
        //wb = Workbook();
        //new_sheet = wb.create_sheet();
        //named_range = NamedRange("test_nr", [(new_sheet, "A1")]);
        //wb.add_named_range(named_range);
        //found_named_range = wb.get_named_range("test_nr");
        //TS_ASSERT_EQUALS(named_range, found_named_range);
    }

    void test_remove_named_range()
    {
        //wb = Workbook();
        //new_sheet = wb.create_sheet();
        //named_range = NamedRange("test_nr", [(new_sheet, "A1")]);
        //wb.add_named_range(named_range);
        //wb.remove_named_range(named_range);
        //named_ranges_list = wb.get_named_ranges();
        //assert named_range not in named_ranges_list;
    }

    void test_add_local_named_range()
    {
        //make_tmpdir();
        //wb = Workbook();
        //new_sheet = wb.create_sheet();
        //named_range = NamedRange("test_nr", [(new_sheet, "A1")]);
        //named_range.scope = new_sheet;
        //wb.add_named_range(named_range);
        //dest_filename = osp.join(TMPDIR, "local_named_range_book.xlsx");
        //wb.save(dest_filename);
        //clean_tmpdir();
    }

    void test_write_regular_date2()
    {
        //make_tmpdir();
        //today = datetime.datetime(2010, 1, 18, 14, 15, 20, 1600);

        //book = Workbook();
        //sheet = book.get_active_sheet();
        //sheet.cell("A1").value = today;
        //dest_filename = osp.join(TMPDIR, "date_read_write_issue.xlsx");
        //book.save(dest_filename);

        //test_book = load_workbook(dest_filename);
        //test_sheet = test_book.get_active_sheet();

        //TS_ASSERT_EQUALS(test_sheet.cell("A1").value, today);
        //clean_tmpdir();
    }

    void test_write_regular_float2()
    {
        //make_tmpdir();
        //float float_value = 1.0 / 3.0;
        //xlnt::workbook book;
        //xlnt::worksheet sheet = book.get_active_sheet();
        //sheet.cell("A1") = float_value;
        //std::string dest_filename = osp.join(TMPDIR, "float_read_write_issue.xlsx");
        //book.save(dest_filename);

        //xlnt::workbook test_book;
        //test_book.load(dest_filename);
        //xlnt::worksheet test_sheet = test_book.get_active_sheet();

        //TS_ASSERT_EQUALS(test_sheet.cell("A1"), float_value);
        //clean_tmpdir();
    }

     //       @raises(UnicodeDecodeError)
    void test_bad_encoding()
    {
        /*pound = chr(163);
        test_string = ("Compound Value (" + pound + ")").encode("latin1");

        utf_book = Workbook();
        utf_sheet = utf_book.get_active_sheet();
        utf_sheet.cell("A1").value = test_string;*/
    }

    void test_good_encoding()
    {
        //char pound = 163;
        //std::string test_string = ("Compound Value (" + std::string(1, pound) + ")").encode("latin1");

        //xlnt::workbook lat_book("latin1");
        //xlnt::worksheet lat_sheet = lat_book.get_active_sheet();
        //lat_sheet.cell("A1") = test_string;
    }
};
