#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/workbook.h>
#include <xlnt/worksheet.h>

class WorkbookTestSuite : public CxxTest::TestSuite
{
public:
    WorkbookTestSuite()
    {

    }

    void test_get_active_sheet()
    {
        //wb = Workbook();
        //active_sheet = wb.get_active_sheet();
        //TS_ASSERT_EQUALS(active_sheet, wb.worksheets[0]);
    }

    void test_create_sheet()
    {
        //wb = Workbook();
        //new_sheet = wb.create_sheet(0);
        //TS_ASSERT_EQUALS(new_sheet, wb.worksheets[0]);
    }

    void test_create_sheet_with_name()
    {
        //wb = Workbook();
        //new_sheet = wb.create_sheet(0, title = "LikeThisName");
        //TS_ASSERT_EQUALS(new_sheet, wb.worksheets[0]);
    }

    void test_add_correct_sheet()
    {
        //wb = Workbook();
        //new_sheet = wb.create_sheet(0);
        //wb.add_sheet(new_sheet);
        //TS_ASSERT_EQUALS(new_sheet, wb.worksheets[2]);
    }

    void test_add_incorrect_sheet()
    {
        //wb = Workbook();
        //wb.add_sheet("Test");
    }

    void test_create_sheet_readonly()
    {
        //wb = Workbook();
        //wb._set_optimized_read();
        //wb.create_sheet();
    }

    void test_remove_sheet()
    {
        //wb = Workbook();
        //new_sheet = wb.create_sheet(0);
        //wb.remove_sheet(new_sheet);
        //assert new_sheet not in wb.worksheets;
    }

    void test_get_sheet_by_name()
    {
        //wb = Workbook();
        //new_sheet = wb.create_sheet();
        //title = "my sheet";
        //new_sheet.title = title;
        //found_sheet = wb.get_sheet_by_name(title);
        //TS_ASSERT_EQUALS(new_sheet, found_sheet);
    }

    void test_get_index2()
    {
        //wb = Workbook();
        //new_sheet = wb.create_sheet(0);
        //sheet_index = wb.get_index(new_sheet);
        //TS_ASSERT_EQUALS(sheet_index, 0);
    }

    void test_get_sheet_names()
    {
        //wb = Workbook();
        //names = ["Sheet", "Sheet1", "Sheet2", "Sheet3", "Sheet4", "Sheet5"];
        //for(auto count : range(5))
        //{
        //    wb.create_sheet(0)
        //        actual_names = wb.get_sheet_names()
        //        TS_ASSERT_EQUALS(sorted(actual_names), sorted(names))
        //}
    }

    void test_get_named_ranges2()
    {
        /*wb = Workbook();
        TS_ASSERT_EQUALS(wb.get_named_ranges(), wb._named_ranges);*/
    }
    void test_get_active_sheet2()
    {
        //wb = Workbook();
        //active_sheet = wb.get_active_sheet();
        //TS_ASSERT_EQUALS(active_sheet, wb.worksheets[0]);
    }

    void test_create_sheet2()
    {
        /*wb = Workbook();
        new_sheet = wb.create_sheet(0);
        TS_ASSERT_EQUALS(new_sheet, wb.worksheets[0]);*/
    }

    void test_create_sheet_with_name2()
    {
        /*wb = Workbook();
        new_sheet = wb.create_sheet(0, title = "LikeThisName");
        TS_ASSERT_EQUALS(new_sheet, wb.worksheets[0]);*/
    }

    void test_add_correct_sheet2()
    {
        /*wb = Workbook();
        new_sheet = wb.create_sheet(0);
        wb.add_sheet(new_sheet);
        TS_ASSERT_EQUALS(new_sheet, wb.worksheets[2]);*/
    }

        //@raises(AssertionError)
    void test_add_incorrect_sheet2()
    {
        /*wb = Workbook();
        wb.add_sheet("Test");*/
    }

        //@raises(ReadOnlyWorkbookException)
    void test_create_sheet_readonly2()
    {
        /*wb = Workbook();
        wb._set_optimized_read();
        wb.create_sheet();*/
    }

    void test_remove_sheet2()
    {
        /*wb = Workbook();
        new_sheet = wb.create_sheet(0);
        wb.remove_sheet(new_sheet);
        assert new_sheet not in wb.worksheets;*/
    }

    void test_get_sheet_by_name2()
    {
        /*wb = Workbook();
        new_sheet = wb.create_sheet();
        title = "my sheet";
        new_sheet.title = title;
        found_sheet = wb.get_sheet_by_name(title);
        TS_ASSERT_EQUALS(new_sheet, found_sheet);*/
    }

    void test_get_index()
    {
        /*wb = Workbook();
        new_sheet = wb.create_sheet(0);
        sheet_index = wb.get_index(new_sheet);
        TS_ASSERT_EQUALS(sheet_index, 0);*/
    }

    void test_get_sheet_names2()
    {
        /*wb = Workbook();
        names = ["Sheet", "Sheet1", "Sheet2", "Sheet3", "Sheet4", "Sheet5"];
        for(auto count in range(5))
        {
            wb.create_sheet(0)
                actual_names = wb.get_sheet_names()
                TS_ASSERT_EQUALS(sorted(actual_names), sorted(names))
        }*/
    }

    void test_get_named_ranges()
    {
       /* wb = Workbook();
        TS_ASSERT_EQUALS(wb.get_named_ranges(), wb._named_ranges);*/
    }

    void test_add_named_range()
    {
        /*wb = Workbook();
        new_sheet = wb.create_sheet();
        named_range = NamedRange("test_nr", [(new_sheet, "A1")]);
        wb.add_named_range(named_range);
        named_ranges_list = wb.get_named_ranges();
        assert named_range in named_ranges_list;*/
    }

    void test_get_named_range2()
    {
        /*wb = Workbook();
        new_sheet = wb.create_sheet();
        named_range = NamedRange("test_nr", [(new_sheet, "A1")]);
        wb.add_named_range(named_range);
        found_named_range = wb.get_named_range("test_nr");
        TS_ASSERT_EQUALS(named_range, found_named_range);*/
    }

    void test_remove_named_range2()
    {
        /*wb = Workbook();
        new_sheet = wb.create_sheet();
        named_range = NamedRange("test_nr", [(new_sheet, "A1")]);
        wb.add_named_range(named_range);
        wb.remove_named_range(named_range);
        named_ranges_list = wb.get_named_ranges();
        assert named_range not in named_ranges_list;*/
    }

            //@with_setup(setup = make_tmpdir, teardown = clean_tmpdir)
    void test_add_local_named_range2()
    {
        /*wb = Workbook();
        new_sheet = wb.create_sheet();
        named_range = NamedRange("test_nr", [(new_sheet, "A1")]);
        named_range.scope = new_sheet;
        wb.add_named_range(named_range);
        dest_filename = osp.join(TMPDIR, "local_named_range_book.xlsx");
        wb.save(dest_filename);*/
    }

            //@with_setup(setup = make_tmpdir, teardown = clean_tmpdir)
    void test_write_regular_date()
    {
        /*today = datetime.datetime(2010, 1, 18, 14, 15, 20, 1600);

        book = Workbook();
        sheet = book.get_active_sheet();
        sheet.cell("A1").value = today;
        dest_filename = osp.join(TMPDIR, "date_read_write_issue.xlsx");
        book.save(dest_filename);

        test_book = load_workbook(dest_filename);
        test_sheet = test_book.get_active_sheet();

        TS_ASSERT_EQUALS(test_sheet.cell("A1").value, today);*/
    }

     //       @with_setup(setup = make_tmpdir, teardown = clean_tmpdir)
    void test_write_regular_float()
    {
        /*float_value = 1.0 / 3.0;
        book = Workbook();
        sheet = book.get_active_sheet();
        sheet.cell("A1").value = float_value;
        dest_filename = osp.join(TMPDIR, "float_read_write_issue.xlsx");
        book.save(dest_filename);

        test_book = load_workbook(dest_filename);
        test_sheet = test_book.get_active_sheet();

        TS_ASSERT_EQUALS(test_sheet.cell("A1").value, float_value);*/
    }

            //@raises(UnicodeDecodeError)
    void test_bad_encoding2()
    {
        /*pound = chr(163);
        test_string = ("Compound Value (" + pound + ")").encode("latin1");

        utf_book = Workbook();
        utf_sheet = utf_book.get_active_sheet();
        utf_sheet.cell("A1").value = test_string;*/
    }

    void test_good_encoding2()
    {
        /*pound = chr(163);
        test_string = ("Compound Value (" + pound + ")").encode("latin1");

        lat_book = Workbook(encoding = "latin1");
        lat_sheet = lat_book.get_active_sheet();
        lat_sheet.cell("A1").value = test_string;*/
    }

    void test_add_named_range2()
    {
        /*wb = Workbook();
        new_sheet = wb.create_sheet();
        named_range = NamedRange("test_nr", [(new_sheet, "A1")]);
        wb.add_named_range(named_range);
        named_ranges_list = wb.get_named_ranges();
        assert named_range in named_ranges_list;*/
    }

    void test_get_named_range()
    {
        /*wb = Workbook();
        new_sheet = wb.create_sheet();
        named_range = NamedRange("test_nr", [(new_sheet, "A1")]);
        wb.add_named_range(named_range);
        found_named_range = wb.get_named_range("test_nr");
        TS_ASSERT_EQUALS(named_range, found_named_range);*/
    }

    void test_remove_named_range()
    {
        /*wb = Workbook();
        new_sheet = wb.create_sheet();
        named_range = NamedRange("test_nr", [(new_sheet, "A1")]);
        wb.add_named_range(named_range);
        wb.remove_named_range(named_range);
        named_ranges_list = wb.get_named_ranges();
        assert named_range not in named_ranges_list;*/
    }

            //@with_setup(setup = make_tmpdir, teardown = clean_tmpdir)
    void test_add_local_named_range()
    {
        /*wb = Workbook();
        new_sheet = wb.create_sheet();
        named_range = NamedRange("test_nr", [(new_sheet, "A1")]);
        named_range.scope = new_sheet;
        wb.add_named_range(named_range);
        dest_filename = osp.join(TMPDIR, "local_named_range_book.xlsx");
        wb.save(dest_filename);*/
    }


     //       @with_setup(setup = make_tmpdir, teardown = clean_tmpdir)
    void test_write_regular_date2()
    {
        /*today = datetime.datetime(2010, 1, 18, 14, 15, 20, 1600);

        book = Workbook();
        sheet = book.get_active_sheet();
        sheet.cell("A1").value = today;
        dest_filename = osp.join(TMPDIR, "date_read_write_issue.xlsx");
        book.save(dest_filename);

        test_book = load_workbook(dest_filename);
        test_sheet = test_book.get_active_sheet();

        TS_ASSERT_EQUALS(test_sheet.cell("A1").value, today);*/
    }

     //       @with_setup(setup = make_tmpdir, teardown = clean_tmpdir)
    void test_write_regular_float2()
    {
        /*float_value = 1.0 / 3.0;
        book = Workbook();
        sheet = book.get_active_sheet();
        sheet.cell("A1").value = float_value;
        dest_filename = osp.join(TMPDIR, "float_read_write_issue.xlsx");
        book.save(dest_filename);

        test_book = load_workbook(dest_filename);
        test_sheet = test_book.get_active_sheet();

        TS_ASSERT_EQUALS(test_sheet.cell("A1").value, float_value);*/
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
        /*pound = chr(163);
        test_string = ("Compound Value (" + pound + ")").encode("latin1");

        lat_book = Workbook(encoding = "latin1");
        lat_sheet = lat_book.get_active_sheet();
        lat_sheet.cell("A1").value = test_string;*/
    }
};
