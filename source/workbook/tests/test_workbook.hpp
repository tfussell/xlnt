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
        TS_ASSERT_THROWS(wb.create_sheet(std::string(40, 'a')), std::runtime_error);
        TS_ASSERT_THROWS(wb.create_sheet("a:a"), std::runtime_error);
        TS_ASSERT_EQUALS(wb.create_sheet("LikeThisName").get_title(), "LikeThisName1");
    }

    void test_create_sheet_with_name_at_index()
    {
        xlnt::workbook wb;
        TS_ASSERT_EQUALS(wb.get_sheet_by_index(0).get_title(), "Sheet");
        wb.create_sheet(0, "LikeThisName");
        TS_ASSERT_EQUALS(wb.get_sheet_by_index(0).get_title(), "LikeThisName");
        TS_ASSERT_EQUALS(wb.get_sheet_by_index(1).get_title(), "Sheet");
    }
    
    void test_add_correct_sheet()
    {
        xlnt::workbook wb;
        auto new_sheet = wb.create_sheet();
        new_sheet.get_cell("A6").set_value(1.498);
        wb.copy_sheet(new_sheet);
        TS_ASSERT(wb[1].compare(wb[2], false));
        wb.create_sheet().get_cell("A6").set_value(1.497);
        TS_ASSERT(!wb[1].compare(wb[3], false));
    }

    void test_add_sheet_from_other_workbook()
    {
        xlnt::workbook wb1, wb2;
        auto new_sheet = wb1.get_active_sheet();
        TS_ASSERT_THROWS(wb2.copy_sheet(new_sheet), xlnt::value_error);
        TS_ASSERT_THROWS(wb2.get_index(new_sheet), std::runtime_error);
    }

    void test_add_sheet_at_index()
    {
        xlnt::workbook wb;
        auto ws = wb.get_active_sheet();
        ws.get_cell("B3").set_value(2);
        ws.set_title("Active");
        wb.copy_sheet(ws, 0);
        TS_ASSERT_EQUALS(wb.get_sheet_names().at(0), "Sheet");
        TS_ASSERT_EQUALS(wb.get_sheet_by_index(0).get_cell("B3").get_value<int>(), 2);
        TS_ASSERT_EQUALS(wb.get_sheet_names().at(1), "Active");
        TS_ASSERT_EQUALS(wb.get_sheet_by_index(1).get_cell("B3").get_value<int>(), 2);
    }
    
    void test_create_sheet_readonly()
    {
        xlnt::workbook wb;
        wb.set_read_only(true);
        TS_ASSERT_THROWS(wb.create_sheet(), xlnt::read_only_workbook_error);
    }
    
    void test_remove_sheet()
    {
        xlnt::workbook wb, wb2;
        auto new_sheet = wb.create_sheet(0);
		new_sheet.set_title("removed");
        wb.remove_sheet(new_sheet);
        TS_ASSERT(!wb.contains("removed"));
        TS_ASSERT_THROWS(wb.remove_sheet(wb2.get_active_sheet()), std::runtime_error);
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
        wb.create_sheet().set_title("1");
        wb.create_sheet().set_title("2");

        auto sheet_index = wb.get_index(wb.get_sheet_by_name("1"));
        TS_ASSERT_EQUALS(sheet_index, 1);

        sheet_index = wb.get_index(wb.get_sheet_by_name("2"));
        TS_ASSERT_EQUALS(sheet_index, 2);
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
        xlnt::workbook wb, wb2;
        auto new_sheet = wb.create_sheet();
        wb.create_named_range("test_nr", new_sheet, "A1");
        TS_ASSERT(new_sheet.has_named_range("test_nr"));
        TS_ASSERT(wb.has_named_range("test_nr"));
        TS_ASSERT_THROWS(wb2.create_named_range("test_nr", new_sheet, "A1"), std::runtime_error);
    }
    
    void test_get_named_range()
    {
        xlnt::workbook wb;
        auto new_sheet = wb.create_sheet();
        wb.create_named_range("test_nr", new_sheet, "A1");
        auto found_range = wb.get_named_range("test_nr");
        auto expected_range = new_sheet.get_range("A1");
        TS_ASSERT_EQUALS(expected_range, found_range);
        TS_ASSERT_THROWS(wb.get_named_range("test_nr2"), std::runtime_error);
    }

    void test_remove_named_range()
    {
        xlnt::workbook wb;
        auto new_sheet = wb.create_sheet();
        wb.create_named_range("test_nr", new_sheet, "A1");
        wb.remove_named_range("test_nr");
        TS_ASSERT(!new_sheet.has_named_range("test_nr"));
        TS_ASSERT(!wb.has_named_range("test_nr"));
        TS_ASSERT_THROWS(wb.remove_named_range("test_nr2"), std::runtime_error);
    }

    void test_write_regular_date()
    {
        const xlnt::datetime today(2010, 1, 18, 14, 15, 20, 1600);
        
        xlnt::workbook book;
        auto sheet = book.get_active_sheet();
        sheet.get_cell("A1").set_value(today);
        temporary_file temp_file;
        book.save(temp_file.get_filename());

        xlnt::workbook test_book;
        test_book.load(temp_file.get_filename());
        auto test_sheet = test_book.get_active_sheet();

        TS_ASSERT_EQUALS(test_sheet.get_cell("A1").get_value<xlnt::datetime>(), today);
    }

    void test_write_regular_float()
    {
        long double float_value = 1.0L / 3.0L;
        
        xlnt::workbook book;
        auto sheet = book.get_active_sheet();
        sheet.get_cell("A1").set_value(float_value);
        temporary_file temp_file;
        book.save(temp_file.get_filename());

        xlnt::workbook test_book;
        test_book.load(temp_file.get_filename());
        auto test_sheet = test_book.get_active_sheet();

        TS_ASSERT_EQUALS(test_sheet.get_cell("A1").get_value<long double>(), float_value);
    }

    void test_post_increment_iterator()
    {
        xlnt::workbook wb;

        wb.create_sheet("Sheet1");
        wb.create_sheet("Sheet2");

        auto iter = wb.begin();

        TS_ASSERT_EQUALS((*(iter++)).get_title(), "Sheet");
        TS_ASSERT_EQUALS((*(iter++)).get_title(), "Sheet1");
        TS_ASSERT_EQUALS((*(iter++)).get_title(), "Sheet2");
        TS_ASSERT_EQUALS(iter, wb.end());

        const auto wb_const = wb;

        auto const_iter = wb_const.begin();

        TS_ASSERT_EQUALS((*(const_iter++)).get_title(), "Sheet");
        TS_ASSERT_EQUALS((*(const_iter++)).get_title(), "Sheet1");
        TS_ASSERT_EQUALS((*(const_iter++)).get_title(), "Sheet2");
        TS_ASSERT_EQUALS(const_iter, wb_const.end());
    }

    void test_copy_iterator()
    {
        xlnt::workbook wb;

        wb.create_sheet("Sheet1");
        wb.create_sheet("Sheet2");

        auto iter = wb.begin();

        iter++;
        TS_ASSERT_EQUALS((*iter).get_title(), "Sheet1");

        auto copy = wb.begin();
        copy = iter;
        TS_ASSERT_EQUALS((*iter).get_title(), "Sheet1");
        TS_ASSERT_EQUALS(iter, copy);

        iter++;
        TS_ASSERT_EQUALS((*iter).get_title(), "Sheet2");
        TS_ASSERT_DIFFERS(iter, copy);

        copy++;
        TS_ASSERT_EQUALS((*iter).get_title(), "Sheet2");
        TS_ASSERT_EQUALS(iter, copy);
    }
    
    void test_manifest()
    {
        xlnt::default_type d;
        TS_ASSERT(d.get_content_type().empty());
        TS_ASSERT(d.get_extension().empty());

        xlnt::override_type o;
        TS_ASSERT(o.get_content_type().empty());
        TS_ASSERT(o.get_part_name().empty());
    }

    void test_get_bad_relationship()
    {
        xlnt::workbook wb;
        TS_ASSERT_THROWS(wb.get_relationship("bad"), std::runtime_error);
    }

    void test_memory()
    {
        xlnt::workbook wb, wb2;
        wb.get_active_sheet().set_title("swap");
        std::swap(wb, wb2);
        TS_ASSERT_EQUALS(wb.get_active_sheet().get_title(), "Sheet");
        TS_ASSERT_EQUALS(wb2.get_active_sheet().get_title(), "swap");
        wb = wb2;
        TS_ASSERT_EQUALS(wb.get_active_sheet().get_title(), "swap");
    }

    void test_clear()
    {
        xlnt::workbook wb;
        xlnt::style s = wb.create_style("s");
        wb.get_active_sheet().get_cell("B2").set_value("B2");
        wb.get_active_sheet().get_cell("B2").set_style(s);
        TS_ASSERT(wb.get_active_sheet().get_cell("B2").has_style());
        wb.clear_styles();
        TS_ASSERT(!wb.get_active_sheet().get_cell("B2").has_style());
        xlnt::format format;
        xlnt::font font;
        font.set_size(41);
        format.set_font(font);
        wb.get_active_sheet().get_cell("B2").set_format(format);
        TS_ASSERT(wb.get_active_sheet().get_cell("B2").has_format());
        wb.clear_formats();
        TS_ASSERT(!wb.get_active_sheet().get_cell("B2").has_format());
        wb.clear();
        TS_ASSERT(wb.get_sheet_names().empty());
    }
};
