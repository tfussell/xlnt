#pragma once

#include <algorithm>
#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/xlnt.hpp>
#include "helpers/temporary_file.hpp"

//checked with 52f25d6

class test_workbook : public CxxTest::TestSuite
{
public:
    void test_get_active_sheet()
    {
        xlnt::workbook wb;
        TS_ASSERT_EQUALS(wb.get_active_sheet(), wb[0]);
    }

    void test_create_sheet()
    {
        xlnt::workbook wb;
        auto new_sheet = wb.create_sheet();
        auto last = std::distance(wb.begin(), wb.end()) - 1;
        TS_ASSERT_EQUALS(new_sheet, wb[last]);
    }

    void test_create_sheet_with_name()
    {
        xlnt::workbook wb;
        auto new_sheet = wb.create_sheet("LikeThisName");
        auto last = std::distance(wb.begin(), wb.end()) - 1;
        TS_ASSERT_EQUALS(new_sheet, wb[last]);
    }
    
    void test_add_correct_sheet()
    {
        xlnt::workbook wb;
        auto new_sheet = wb.create_sheet();
        wb.add_sheet(new_sheet);
        TS_ASSERT(wb[1].compare(wb[2], false));
    }
    
    // void test_add_sheetname() {} unnecessary

    void test_add_sheet_from_other_workbook()
    {
        xlnt::workbook wb1, wb2;
        auto new_sheet = wb1.get_active_sheet();
        TS_ASSERT_THROWS(wb2.add_sheet(new_sheet), xlnt::value_error);
    }
    
    void test_create_sheet_readonly()
    {
        xlnt::workbook wb;
        wb.set_read_only(true);
        TS_ASSERT_THROWS(wb.create_sheet(), xlnt::read_only_workbook_exception);
    }
    
    void test_remove_sheet()
    {
        xlnt::workbook wb;
        auto new_sheet = wb.create_sheet(0);
        wb.remove_sheet(new_sheet);
        TS_ASSERT(std::find(wb.begin(), wb.end(), new_sheet) == wb.end());
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
    
    void test_get_sheet_by_name_const()
    {
        xlnt::workbook wb;
        auto new_sheet = wb.create_sheet();
        std::string title = "my sheet";
        new_sheet.set_title(title);
        const xlnt::workbook& wbconst = wb;
        auto found_sheet = wbconst.get_sheet_by_name(title);
        TS_ASSERT_EQUALS(new_sheet, found_sheet);
    }

    void test_index_operator() // test_getitem
    {
        xlnt::workbook wb;
        TS_ASSERT_THROWS_NOTHING(wb["Sheet"]);
        TS_ASSERT_THROWS(wb["NotThere"], xlnt::key_error);
    }
    
    // void test_delitem() {} doesn't make sense in c++
    
    void test_contains()
    {
        xlnt::workbook wb;
        TS_ASSERT(wb.contains("Sheet"));
        TS_ASSERT(!wb.contains("NotThere"));
    }
    
    void test_iter()
    {
        xlnt::workbook wb;
        
        for(auto ws : wb)
        {
            TS_ASSERT_EQUALS(ws.get_title(), "Sheet");
        }
    }
    
    void test_get_index()
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
        
        std::vector<std::string> names = {"Sheet", "Sheet1", "Sheet2", "Sheet3", "Sheet4", "Sheet5"};
        
        for(std::size_t count = 0; count < names.size(); count++)
        {
            wb.create_sheet();
        }
        
        auto actual_names = wb.get_sheet_names();
        
        std::sort(actual_names.begin(), actual_names.end());
        std::sort(names.begin(), names.end());
        
        for(std::size_t count = 0; count < names.size(); count++)
        {
            TS_ASSERT_EQUALS(actual_names[count], names[count]);
        }
    }
    
    void test_add_named_range()
    {
        xlnt::workbook wb;
        auto new_sheet = wb.create_sheet();
        wb.create_named_range("test_nr", new_sheet, "A1");
        TS_ASSERT(new_sheet.has_named_range("test_nr"));
        TS_ASSERT(wb.has_named_range("test_nr"));
    }
    
    void test_get_named_range()
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

    void test_write_regular_date()
    {
        const xlnt::datetime today(2010, 1, 18, 14, 15, 20, 1600);
        
        xlnt::workbook book;
        auto sheet = book.get_active_sheet();
        sheet.get_cell("A1").set_value(today);
        TemporaryFile temp_file;
        book.save(temp_file.GetFilename());

        xlnt::workbook test_book;
        test_book.load(temp_file.GetFilename());
        auto test_sheet = test_book.get_active_sheet();

        TS_ASSERT_EQUALS(test_sheet.get_cell("A1").get_value<xlnt::datetime>(), today);
    }

    void test_write_regular_float()
    {
        long double float_value = 1.0L / 3.0L;
        
        xlnt::workbook book;
        auto sheet = book.get_active_sheet();
        sheet.get_cell("A1").set_value(float_value);
        TemporaryFile temp_file;
        book.save(temp_file.GetFilename());

        xlnt::workbook test_book;
        test_book.load(temp_file.GetFilename());
        auto test_sheet = test_book.get_active_sheet();

        TS_ASSERT_EQUALS(test_sheet.get_cell("A1").get_value<long double>(), float_value);
    }
    
    // void test_add_invalid_worksheet_class_instance() {} not needed in c++
};
