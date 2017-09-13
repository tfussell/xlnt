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

#include <iostream>
#include <stdexcept>

#include <helpers/test_suite.hpp>
#include <helpers/xml_helper.hpp>
#include <helpers/assertions.hpp>

class helper_test_suite : public test_suite
{
public:
    helper_test_suite()
    {
        register_test(test_bootstrap_asserts);
        register_test(test_other_asserts);
        register_test(test_compare);
    }

    // bootstrap asserts are used by other assert tests so they need to be tested first
    // without using themselves
    void test_bootstrap_asserts()
    {
        auto function_that_throws = []()
        {
            throw xlnt::exception("test");
        };

        auto function_that_doesnt_throw = []()
        {
            return std::string("abc");
        };

        // 1. function that throws no exception shouldn't cause exception to be thrown in xlnt_assert_throws_nothing
        try
        {
            xlnt_assert_throws_nothing(function_that_doesnt_throw());
            // good
        }
        catch (...)
        {
            throw xlnt::exception("test failed");
        }


        // 2. function that throws an exception should cause exception to be thrown in xlnt_assert_throws_nothing
        try
        {
            xlnt_assert_throws_nothing(function_that_throws());
        }
        catch (std::runtime_error)
        {
            // good
        }
        catch (...)
        {
            throw xlnt::exception("test failed");
        }


        // 3. function that doesn't throw an exception should cause exception to be thrown in xlnt_assert_throws
        try
        {
            xlnt_assert_throws(function_that_throws(), std::runtime_error);
            // good
        }
        catch (...)
        {
            throw xlnt::exception("test failed");
        }


        // 4. function that doesn't throw an exception should cause exception to be thrown in xlnt_assert_throws
        try
        {
            xlnt_assert_throws(function_that_doesnt_throw(), std::runtime_error);
            throw xlnt::exception("test failed");
        }
        catch (xlnt::exception)
        {
            // good
        }
        catch (...)
        {
            throw xlnt::exception("test failed");
        }


        // 5. function that doesn't throw an exception should cause exception to be thrown in xlnt_assert_throws
        try
        {
            xlnt_assert_throws(function_that_throws(), xlnt::exception);
            throw xlnt::exception("test failed");
        }
        catch (xlnt::exception)
        {
            // good
        }
        catch (...)
        {
            throw xlnt::exception("test failed");
        }
    }

    void test_other_asserts()
    {
        xlnt_assert_throws_nothing(xlnt_assert(true));
        xlnt_assert_throws(xlnt_assert(false), xlnt::exception);

        xlnt_assert_throws_nothing(xlnt_assert_equals(1, 1));
        xlnt_assert_throws(xlnt_assert_equals(1, 2), xlnt::exception);

        xlnt_assert_throws_nothing(xlnt_assert_differs(1, 2));
        xlnt_assert_throws(xlnt_assert_differs(1, 1), xlnt::exception);

        xlnt_assert_throws_nothing(xlnt_assert_delta(1.0, 1.05, 0.06));
        xlnt_assert_throws(xlnt_assert_delta(1.0, 1.05, 0.04), xlnt::exception);
    }

    void test_compare()
    {
        xlnt_assert(!xml_helper::compare_xml_exact("<a/>", "<b/>", true));
        xlnt_assert(!xml_helper::compare_xml_exact("<a/>", "<a b=\"4\"/>", true));
        xlnt_assert(!xml_helper::compare_xml_exact("<a b=\"3\"/>", "<a/>", true));
        xlnt_assert(!xml_helper::compare_xml_exact("<a c=\"4\"/>", "<a b=\"4\"/>", true));
        xlnt_assert(!xml_helper::compare_xml_exact("<a b=\"3\"/>", "<a b=\"4\"/>", true));
        xlnt_assert(!xml_helper::compare_xml_exact("<a>text</a>", "<a>txet</a>", true));
        xlnt_assert(!xml_helper::compare_xml_exact("<a>text</a>", "<a><b>txet</b></a>", true));
        xlnt_assert(xml_helper::compare_xml_exact("<a/>", "<a> </a>", true));
        xlnt_assert(xml_helper::compare_xml_exact("<a b=\"3\"/>", "<a b=\"3\"></a>", true));
        xlnt_assert(xml_helper::compare_xml_exact("<a>text</a>", "<a>text</a>", true));
    }
};
