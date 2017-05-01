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

#include <iostream>

#include <cell/cell_test_suite.hpp>
#include <cell/index_types_test_suite.hpp>
#include <cell/rich_text_test_suite.hpp>

#include <styles/alignment_test_suite.hpp>
#include <styles/color_test_suite.hpp>
#include <styles/fill_test_suite.hpp>
#include <styles/number_format_test_suite.hpp>

#include <utils/datetime_test_suite.hpp>
#include <utils/path_test_suite.hpp>
#include <utils/helper_test_suite.hpp>
#include <utils/timedelta_test_suite.hpp>

#include <workbook/named_range_test_suite.hpp>
#include <workbook/serialization_test_suite.hpp>
#include <workbook/workbook_test_suite.hpp>

#include <worksheet/page_setup_test_suite.hpp>
#include <worksheet/range_test_suite.hpp>
#include <worksheet/worksheet_test_suite.hpp>

#include <detail/cryptography/compound_document.hpp>

test_status overall_status;

template<typename T>
void run_tests()
{
    auto status = T{}.go();

    overall_status.tests_run += status.tests_run;
    overall_status.tests_passed += status.tests_passed;
    overall_status.tests_failed += status.tests_failed;

    std::copy(status.failures.begin(), status.failures.end(), std::back_inserter(overall_status.failures));
}

void print_summary()
{
    std::cout << std::endl;

    for (auto failure : overall_status.failures)
    {
        std::cout << failure << " failed" << std::endl;
    }
}

int main()
{
    // cell
    run_tests<cell_test_suite>();
    run_tests<index_types_test_suite>();
    run_tests<rich_text_test_suite>();

    // styles
    run_tests<alignment_test_suite>();
    run_tests<color_test_suite>();
    run_tests<fill_test_suite>();
    run_tests<number_format_test_suite>();

    // utils
    run_tests<datetime_test_suite>();
    run_tests<path_test_suite>();
    run_tests<helper_test_suite>();
    run_tests<timedelta_test_suite>();

    // workbook
    run_tests<named_range_test_suite>();
    run_tests<serialization_test_suite>();
    run_tests<workbook_test_suite>();

    // worksheet
    run_tests<page_setup_test_suite>();
    run_tests<range_test_suite>();
    run_tests<worksheet_test_suite>();

    print_summary();

    return static_cast<int>(overall_status.tests_failed);
}
