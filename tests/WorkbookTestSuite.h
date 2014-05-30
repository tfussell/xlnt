#pragma once

#include <algorithm>
#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/xlnt.h>

class WorkbookTestSuite : public CxxTest::TestSuite
{
public:
    WorkbookTestSuite()
    {

    }

    void test_get_active_sheet()
    {
        xlnt::workbook wb;
        auto active_sheet = wb.get_active_sheet();
        TS_ASSERT_EQUALS(active_sheet, wb[0]);
    }

    void test_create_sheet()
    {
        xlnt::workbook wb;
        auto new_sheet = wb.create_sheet(0);
        TS_ASSERT_EQUALS(new_sheet, wb[0]);
    }

    void test_create_sheet_with_name()
    {
        xlnt::workbook wb;
        auto new_sheet = wb.create_sheet(0, "LikeThisName");
        TS_ASSERT_EQUALS(new_sheet, wb[0]);
    }

    void test_create_sheet_readonly()
    {
        xlnt::workbook wb(xlnt::optimization::read);
        TS_ASSERT_THROWS_ANYTHING(wb.create_sheet());
    }

    void test_remove_sheet()
    {
        xlnt::workbook wb;
        auto new_sheet = wb.create_sheet(0);
        auto sheet_name = new_sheet.get_title();
        wb.remove_sheet(new_sheet);
        for(auto worksheet : wb)
        {
            TS_ASSERT_DIFFERS(sheet_name, worksheet.get_title());
        }
    }

    void test_get_sheet_by_name()
    {
        xlnt::workbook wb;
        auto new_sheet = wb.create_sheet();
        std::string title = "my sheet";
        new_sheet.set_title(title);
        auto found_sheet = wb.get_sheet_by_name(title);
        TS_ASSERT_EQUALS(new_sheet, found_sheet);
    }
    
    void test_getitem()
    {
        xlnt::workbook wb;
        auto new_sheet = wb.create_sheet();
        std::string title = "my sheet";
        new_sheet.set_title(title);
        auto found_sheet = wb.get_sheet_by_name(title);
        TS_ASSERT_EQUALS(new_sheet, found_sheet);
        
        TS_ASSERT_EQUALS(wb.get_sheet_by_name("NotThere"), nullptr);
    }

    void test_get_index2()
    {
        xlnt::workbook wb;
        auto new_sheet = wb.create_sheet(0);
        auto sheet_index = wb.get_index(new_sheet);
        TS_ASSERT_EQUALS(sheet_index, 0);
    }

    void test_get_sheet_names()
    {
        xlnt::workbook wb;
        wb.clear();
        std::vector<std::string> names = {"NewSheet", "NewSheet1", "NewSheet2", "NewSheet3", "NewSheet4", "NewSheet5"};
        for(int count = 0; count < names.size(); count++)
        {
            wb.create_sheet(names[count]);
        }
        auto actual_names = wb.get_sheet_names();
        std::sort(actual_names.begin(), actual_names.end());
        std::sort(names.begin(), names.end());
        for(int count = 0; count < names.size(); count++)
        {
            TS_ASSERT_EQUALS(actual_names[count], names[count]);
        }
    }

    void test_get_active_sheet2()
    {
        xlnt::workbook wb;
        auto active_sheet = wb.get_active_sheet();
        TS_ASSERT_EQUALS(active_sheet, wb[0]);
    }

    void test_create_sheet2()
    {
        xlnt::workbook wb;
        auto new_sheet = wb.create_sheet(0);
        TS_ASSERT_EQUALS(new_sheet, wb[0]);
    }

    void test_create_sheet_with_name2()
    {
        xlnt::workbook wb;
        auto new_sheet = wb.create_sheet(0, "LikeThisName");
        TS_ASSERT_EQUALS(new_sheet, wb[0]);
    }

    void test_get_sheet_by_name2()
    {
        xlnt::workbook wb;
        auto new_sheet = wb.create_sheet();
        std::string title = "my sheet";
        new_sheet.set_title(title);
        auto found_sheet = wb.get_sheet_by_name(title);
        TS_ASSERT_EQUALS(new_sheet, found_sheet);
    }

    void test_get_index()
    {
        xlnt::workbook wb;
        auto new_sheet = wb.create_sheet(0);
        auto sheet_index = wb.get_index(new_sheet);
        TS_ASSERT_EQUALS(sheet_index, 0);
    }

    void test_add_named_range()
    {
        xlnt::workbook wb;
        auto new_sheet = wb.create_sheet();
        wb.create_named_range("test_nr", new_sheet, "A1");
        TS_ASSERT(new_sheet.has_named_range("test_nr"));
        TS_ASSERT(wb.has_named_range("test_nr"));
    }

    void test_get_named_range2()
    {
        xlnt::workbook wb;
        auto new_sheet = wb.create_sheet();
        wb.create_named_range("test_nr", new_sheet, "A1");
        auto found_range = wb.get_named_range("test_nr");
        auto expected_range = new_sheet.get_range("A1");
        TS_ASSERT_EQUALS(expected_range, found_range);
    }

    void test_remove_named_range()
    {
        xlnt::workbook wb;
        auto new_sheet = wb.create_sheet();
        wb.create_named_range("test_nr", new_sheet, "A1");
        wb.remove_named_range("test_nr");
        TS_ASSERT(!new_sheet.has_named_range("test_nr"));
        TS_ASSERT(!wb.has_named_range("test_nr"));
    }

    void test_add_local_named_range()
    {
        TemporaryFile temp_file;
        xlnt::workbook wb;
        auto new_sheet = wb.create_sheet();
        wb.create_named_range("test_nr", new_sheet, "A1");
        auto dest_filename = temp_file.GetFilename();
        wb.save(dest_filename);
    }

    void test_write_regular_date()
    {
        /*make_tmpdir();
        auto today = datetime.datetime(2010, 1, 18, 14, 15, 20, 1600);

        xlnt::workbook wb;
        auto sheet = book.get_active_sheet();
        sheet.cell("A1") = today;
        dest_filename = osp.join(TMPDIR, "date_read_write_issue.xlsx");
        book.save(dest_filename);

        test_book = load_workbook(dest_filename);
        test_sheet = test_book.get_active_sheet();

        TS_ASSERT_EQUALS(test_sheet.cell("A1"), today);
        clean_tmpdir();*/
    }

    void test_write_regular_float()
    {
        TemporaryFile temp_file;
        double float_value = 1.0 / 3.0;
        xlnt::workbook book;
        auto sheet = book.get_active_sheet();
        sheet.get_cell("A1") = float_value;
        auto dest_filename = temp_file.GetFilename();
        book.save(dest_filename);

        xlnt::workbook test_book;
        test_book.load(dest_filename);
        auto test_sheet = test_book.get_active_sheet();

        TS_ASSERT_LESS_THAN_EQUALS(std::stod(test_sheet.get_cell("A1").get_value()) - float_value, 0.001);
    }
};
