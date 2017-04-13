// Copyright (c) 2014-2017 Thomas Fussell
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

#pragma once

#include <algorithm>
#include <iostream>

#include <helpers/temporary_file.hpp>
#include <helpers/test_suite.hpp>
#include <xlnt/xlnt.hpp>

class workbook_test_suite : public test_suite
{
public:
    workbook_test_suite()
    {
        register_test(test_active_sheet);
        register_test(test_create_sheet);
        register_test(test_add_correct_sheet);
        register_test(test_add_sheet_from_other_workbook);
        register_test(test_add_sheet_at_index);
        register_test(test_get_sheet_by_title);
        register_test(test_get_sheet_by_title_const);
        register_test(test_index_operator);
        register_test(test_contains);
        register_test(test_iter);
        register_test(test_get_index);
        register_test(test_get_sheet_names);
        register_test(test_add_named_range);
        register_test(test_get_named_range);
        register_test(test_remove_named_range);
        register_test(test_post_increment_iterator);
        register_test(test_copy_iterator);
        register_test(test_manifest);
        register_test(test_memory);
        register_test(test_clear);
        register_test(test_comparison);
    }

    void test_active_sheet()
    {
        xlnt::workbook wb;
        assert_equals(wb.active_sheet(), wb[0]);
    }

    void test_create_sheet()
    {
        xlnt::workbook wb;
        auto new_sheet = wb.create_sheet();
        auto last = std::distance(wb.begin(), wb.end()) - 1;
        assert_equals(new_sheet, wb[last]);
    }

    void test_add_correct_sheet()
    {
        xlnt::workbook wb;
        auto new_sheet = wb.create_sheet();
        new_sheet.cell("A6").value(1.498);
        wb.copy_sheet(new_sheet);
        assert(wb[1].compare(wb[2], false));
        wb.create_sheet().cell("A6").value(1.497);
        assert(!wb[1].compare(wb[3], false));
    }

    void test_add_sheet_from_other_workbook()
    {
        xlnt::workbook wb1, wb2;
        auto new_sheet = wb1.active_sheet();
        assert_throws(wb2.copy_sheet(new_sheet), xlnt::invalid_parameter);
        assert_throws(wb2.index(new_sheet), std::runtime_error);
    }

    void test_add_sheet_at_index()
    {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        ws.cell("B3").value(2);
        ws.title("Active");
        wb.copy_sheet(ws, 0);
        assert_equals(wb.sheet_titles().at(0), "Sheet1");
        assert_equals(wb.sheet_by_index(0).cell("B3").value<int>(), 2);
        assert_equals(wb.sheet_titles().at(1), "Active");
        assert_equals(wb.sheet_by_index(1).cell("B3").value<int>(), 2);
    }
    
    void test_remove_sheet()
    {
        xlnt::workbook wb, wb2;
        auto new_sheet = wb.create_sheet(0);
		new_sheet.title("removed");
        wb.remove_sheet(new_sheet);
        assert(!wb.contains("removed"));
        assert_throws(wb.remove_sheet(wb2.active_sheet()), std::runtime_error);
    }

    void test_get_sheet_by_title()
    {
        xlnt::workbook wb;
        auto new_sheet = wb.create_sheet();
        std::string title = "my sheet";
        new_sheet.title(title);
        auto found_sheet = wb.sheet_by_title(title);
        assert_equals(new_sheet, found_sheet);
        assert_throws(wb.sheet_by_title("error"), xlnt::key_not_found);
        const auto &wb_const = wb;
        assert_throws(wb_const.sheet_by_title("error"), xlnt::key_not_found);
    }
    
    void test_get_sheet_by_title_const()
    {
        xlnt::workbook wb;
        auto new_sheet = wb.create_sheet();
        std::string title = "my sheet";
        new_sheet.title(title);
        const xlnt::workbook& wbconst = wb;
        auto found_sheet = wbconst.sheet_by_title(title);
        assert_equals(new_sheet, found_sheet);
    }

    void test_index_operator() // test_getitem
    {
        xlnt::workbook wb;
        assert_throws_nothing(wb["Sheet1"]);
        assert_throws(wb["NotThere"], xlnt::key_not_found);
    }

    void test_contains()
    {
        xlnt::workbook wb;
        assert(wb.contains("Sheet1"));
        assert(!wb.contains("NotThere"));
    }
    
    void test_iter()
    {
        xlnt::workbook wb;
        
        for(auto ws : wb)
        {
            assert_equals(ws.title(), "Sheet1");
        }
    }
    
    void test_get_index()
    {
        xlnt::workbook wb;
        wb.create_sheet().title("1");
        wb.create_sheet().title("2");

        auto sheet_index = wb.index(wb.sheet_by_title("1"));
        assert_equals(sheet_index, 1);

        sheet_index = wb.index(wb.sheet_by_title("2"));
        assert_equals(sheet_index, 2);
    }

    void test_get_sheet_names()
    {
        xlnt::workbook wb;
        wb.create_sheet().title("test_get_sheet_titles");

		const std::vector<std::string> expected_titles = { "Sheet1", "test_get_sheet_titles" };

		assert_equals(wb.sheet_titles(), expected_titles);
    }
    
    void test_add_named_range()
    {
        xlnt::workbook wb, wb2;
        auto new_sheet = wb.create_sheet();
        wb.create_named_range("test_nr", new_sheet, "A1");
        assert(new_sheet.has_named_range("test_nr"));
        assert(wb.has_named_range("test_nr"));
        assert_throws(wb2.create_named_range("test_nr", new_sheet, "A1"), std::runtime_error);
    }
    
    void test_get_named_range()
    {
        xlnt::workbook wb;
        auto new_sheet = wb.create_sheet();
        wb.create_named_range("test_nr", new_sheet, "A1");
        auto found_range = wb.named_range("test_nr");
        auto expected_range = new_sheet.range("A1");
        assert_equals(expected_range, found_range);
        assert_throws(wb.named_range("test_nr2"), std::runtime_error);
    }

    void test_remove_named_range()
    {
        xlnt::workbook wb;
        auto new_sheet = wb.create_sheet();
        wb.create_named_range("test_nr", new_sheet, "A1");
        wb.remove_named_range("test_nr");
        assert(!new_sheet.has_named_range("test_nr"));
        assert(!wb.has_named_range("test_nr"));
        assert_throws(wb.remove_named_range("test_nr2"), std::runtime_error);
    }

    void test_post_increment_iterator()
    {
        xlnt::workbook wb;

        wb.create_sheet().title("Sheet2");
        wb.create_sheet().title("Sheet3");

        auto iter = wb.begin();

        assert_equals((*(iter++)).title(), "Sheet1");
        assert_equals((*(iter++)).title(), "Sheet2");
        assert_equals((*(iter++)).title(), "Sheet3");
        assert_equals(iter, wb.end());

        const auto wb_const = wb;

        auto const_iter = wb_const.begin();

        assert_equals((*(const_iter++)).title(), "Sheet1");
        assert_equals((*(const_iter++)).title(), "Sheet2");
        assert_equals((*(const_iter++)).title(), "Sheet3");
        assert_equals(const_iter, wb_const.end());
    }

    void test_copy_iterator()
    {
        xlnt::workbook wb;

        wb.create_sheet().title("Sheet2");
        wb.create_sheet().title("Sheet3");

        auto iter = wb.begin();
		assert_equals((*iter).title(), "Sheet1");

        iter++;
        assert_equals((*iter).title(), "Sheet2");

        auto copy = wb.begin();
        copy = iter;
        assert_equals((*iter).title(), "Sheet2");
        assert_equals(iter, copy);

        iter++;
        assert_equals((*iter).title(), "Sheet3");
        assert_differs(iter, copy);

        copy++;
        assert_equals((*iter).title(), "Sheet3");
        assert_equals(iter, copy);
    }
    
    void test_manifest()
    {
		xlnt::manifest m;
        assert(!m.has_default_type("xml"));
        assert_throws(m.default_type("xml"), xlnt::key_not_found);
        assert(!m.has_relationship(xlnt::path("/"), xlnt::relationship_type::office_document));
        assert(m.relationships(xlnt::path("xl/workbook.xml")).empty());
    }

    void test_memory()
    {
        xlnt::workbook wb, wb2;
        wb.active_sheet().title("swap");
        std::swap(wb, wb2);
        assert_equals(wb.active_sheet().title(), "Sheet1");
        assert_equals(wb2.active_sheet().title(), "swap");
        wb = wb2;
        assert_equals(wb.active_sheet().title(), "swap");
    }

    void test_clear()
    {
        xlnt::workbook wb;
        xlnt::style s = wb.create_style("s");
        wb.active_sheet().cell("B2").value("B2");
        wb.active_sheet().cell("B2").style(s);
        assert(wb.active_sheet().cell("B2").has_style());
        wb.clear_styles();
        assert(!wb.active_sheet().cell("B2").has_style());
		xlnt::format format = wb.create_format();
        xlnt::font font;
        font.size(41);
        format.font(font, true);
        wb.active_sheet().cell("B2").format(format);
        assert(wb.active_sheet().cell("B2").has_format());
        wb.clear_formats();
        assert(!wb.active_sheet().cell("B2").has_format());
        wb.clear();
        assert(wb.sheet_titles().empty());
    }

    void test_comparison()
    {
        xlnt::workbook wb, wb2;
        assert(wb == wb);
        assert(!(wb != wb));
        assert(!(wb == wb2));
        assert(wb != wb2);
        
        const auto &wb_const = wb;
        //TODO these aren't tests...
        wb_const.manifest();
        
        assert(wb.has_theme());
        
        wb.create_style("style1");
        wb.style("style1");
        wb_const.style("style1");
    }
};
