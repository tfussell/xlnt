// Copyright (c) 2014-2018 Thomas Fussell
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

#include <iostream>

#include <helpers/test_suite.hpp>

#include <xlnt/cell/cell.hpp>
#include <xlnt/styles/fill.hpp>
#include <xlnt/utils/date.hpp>
#include <xlnt/worksheet/worksheet.hpp>

class fill_test_suite : public test_suite
{
public:
    fill_test_suite()
    {
        register_test(test_properties);
        register_test(test_comparison);
        register_test(test_two_fills);
    }

    void test_properties()
    {
        xlnt::fill fill;
        
        xlnt_assert_equals(fill.type(), xlnt::fill_type::pattern);
        fill = xlnt::fill(xlnt::gradient_fill());
        xlnt_assert_equals(fill.type(), xlnt::fill_type::gradient);
        xlnt_assert_equals(fill.gradient_fill().type(), xlnt::gradient_fill_type::linear);

        xlnt_assert_equals(fill.gradient_fill().left(), 0);
        xlnt_assert_equals(fill.gradient_fill().right(), 0);
        xlnt_assert_equals(fill.gradient_fill().top(), 0);
        xlnt_assert_equals(fill.gradient_fill().bottom(), 0);

        xlnt_assert_throws(fill.pattern_fill(), xlnt::invalid_attribute);

        xlnt_assert_equals(fill.gradient_fill().degree(), 0);
		fill = fill.gradient_fill().degree(1);
        xlnt_assert_equals(fill.gradient_fill().degree(), 1);

        fill = xlnt::pattern_fill().type(xlnt::pattern_fill_type::solid);

        fill = fill.pattern_fill().foreground(xlnt::color::black());
        xlnt_assert(fill.pattern_fill().foreground().is_set());
        xlnt_assert_equals(fill.pattern_fill().foreground().get().rgb().hex_string(),
            xlnt::color::black().rgb().hex_string());

        fill = fill.pattern_fill().background(xlnt::color::green());
        xlnt_assert(fill.pattern_fill().background().is_set());
        xlnt_assert_equals(fill.pattern_fill().background().get().rgb().hex_string(),
            xlnt::color::green().rgb().hex_string());
        
        const auto &const_fill = fill;
        xlnt_assert(const_fill.pattern_fill().foreground().is_set());
        xlnt_assert(const_fill.pattern_fill().background().is_set());
    }
    
    void test_comparison()
    {
        xlnt::fill pattern_fill = xlnt::pattern_fill().type(xlnt::pattern_fill_type::solid);
        xlnt::fill gradient_fill_linear = xlnt::gradient_fill().type(xlnt::gradient_fill_type::linear);
        xlnt::fill gradient_fill_path = xlnt::gradient_fill().type(xlnt::gradient_fill_type::path);

        xlnt_assert_differs(pattern_fill, gradient_fill_linear);
        xlnt_assert_differs(gradient_fill_linear, gradient_fill_path);
        xlnt_assert_differs(gradient_fill_path, pattern_fill);
    }

    void test_two_fills()
    {
        xlnt::workbook wb;
        auto ws1 = wb.active_sheet();

        auto cell1 = ws1.cell("A10");
        auto cell2 = ws1.cell("A11");

        cell1.fill(xlnt::fill::solid(xlnt::color::yellow()));
        cell1.value("Fill Yellow");

        cell2.fill(xlnt::fill::solid(xlnt::color::green()));
        cell2.value(xlnt::date(2010, 7, 13));

        xlnt_assert_equals(cell1.fill(), xlnt::fill::solid(xlnt::color::yellow()));
        xlnt_assert_equals(cell2.fill(), xlnt::fill::solid(xlnt::color::green()));
    }
};
static fill_test_suite x;