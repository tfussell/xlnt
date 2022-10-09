// Copyright (c) 2014-2021 Thomas Fussell
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
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE
//
// @license: http://www.opensource.org/licenses/mit-license.php
// @author: see AUTHORS file

#include <xlnt/utils/numeric.hpp>
#include <helpers/test_suite.hpp>

class numeric_test_suite : public test_suite
{
public:
    numeric_test_suite()
    {
        register_test(test_serialise_number);
        register_test(test_float_equals_zero);
        register_test(test_float_equals_large);
        register_test(test_float_equals_fairness);
        register_test(test_min);
        register_test(test_max);
        register_test(test_abs);
    }

    void test_serialise_number()
    {
        xlnt::detail::number_serialiser serialiser;
        // excel serialises numbers as floating point values with <= 15 digits of precision
        xlnt_assert(serialiser.serialise(1) == "1");
        // trailing zeroes are ignored
        xlnt_assert(serialiser.serialise(1.0) == "1");
        xlnt_assert(serialiser.serialise(1.0f) == "1");
        // one to 1 relation
        xlnt_assert(serialiser.serialise(1.23456) == "1.23456");
        xlnt_assert(serialiser.serialise(1.23456789012345) == "1.23456789012345");
        xlnt_assert(serialiser.serialise(123456.789012345) == "123456.789012345");
        xlnt_assert(serialiser.serialise(1.23456789012345e+67) == "1.23456789012345e+67");
        xlnt_assert(serialiser.serialise(1.23456789012345e-67) == "1.23456789012345e-67");
    }

    void test_float_equals_zero()
    {
        // comparing relatively small numbers (2.3e-6) with 0 will be true by default
        const float comp_val = 2.3e-6f; // about the largest difference allowed by default
        xlnt_assert(0.f != comp_val); // fail because not exactly equal
        xlnt_assert(xlnt::detail::float_equals(0.0, comp_val));
        xlnt_assert(xlnt::detail::float_equals(0.0, -comp_val));
        // fail because diff is out of bounds for fuzzy equality
        xlnt_assert(!xlnt::detail::float_equals(0.0, comp_val + 0.1e-6));
        xlnt_assert(!xlnt::detail::float_equals(0.0, -(comp_val + 0.1e-6)));
        // if the bounds of comparison are too loose, there are two tweakable knobs to tighten the comparison up
        //==========================================================
        // #1: reduce the epsilon_scale (default is 20)
        // This can bring the range down to FLT_EPSILON (scale factor of 1)
        xlnt_assert(!xlnt::detail::float_equals(0.0, comp_val, 10));
        const float closer_comp_val = 1.1e-6f;
        xlnt_assert(xlnt::detail::float_equals(0.0, closer_comp_val, 10));
        xlnt_assert(!xlnt::detail::float_equals(0.0, closer_comp_val + 0.1e-6, 10));
        xlnt_assert(xlnt::detail::float_equals(0.0, -closer_comp_val, 10));
        xlnt_assert(!xlnt::detail::float_equals(0.0, -(closer_comp_val + 0.1e-6), 10));
        //==========================================================
        // #2: specify the epsilon source as a higher precision type (e.g. double)
        // This makes the epsilon range quite significantly less
        xlnt_assert(!xlnt::detail::float_equals<double>(0.0, comp_val));
        xlnt_assert(!xlnt::detail::float_equals<double>(0.0, closer_comp_val));
        const float tiny_comp_val = 4.4e-15f;
        xlnt_assert(xlnt::detail::float_equals<double>(0.0, tiny_comp_val));
        xlnt_assert(!xlnt::detail::float_equals<double>(0.0, tiny_comp_val + 0.1e-15));
        xlnt_assert(xlnt::detail::float_equals<double>(0.0, -tiny_comp_val));
        xlnt_assert(!xlnt::detail::float_equals<double>(0.0, -(tiny_comp_val + 0.1e-15)));
        //==========================================================
        // #3: combine #1 & #2
        // for the tightest default precision, double with a scale of 1
        xlnt_assert(!xlnt::detail::float_equals<double>(0.0, comp_val, 1));
        xlnt_assert(!xlnt::detail::float_equals<double>(0.0, closer_comp_val, 1));
        xlnt_assert(!xlnt::detail::float_equals<double>(0.0, tiny_comp_val, 1));
        const float really_tiny_comp_val = 2.2e-16f; // the limit is +/- std::numeric_limits<double>::epsilon()
        xlnt_assert(xlnt::detail::float_equals<double>(0.0, really_tiny_comp_val, 1));
        xlnt_assert(!xlnt::detail::float_equals<double>(0.0, really_tiny_comp_val + 0.1e-16, 1));
        xlnt_assert(xlnt::detail::float_equals<double>(0.0, -really_tiny_comp_val, 1));
        xlnt_assert(!xlnt::detail::float_equals<double>(0.0, -(really_tiny_comp_val + 0.1e-16), 1));
        //==========================================================
        // in the world of floats, 2.2e-16 is still significantly different to 0.f (smallest representable float is around 1e-38)
        // if comparisons are known to involve extremely small numbers (such that +/- 2.2e-16 is too large a band),
        // a type that specialises std::numeric_limits::epsilon may be passed as the first template parameter
        // the type itself doesn't actually need to have any behaviour as it is only used as the source for epsilon
        // struct super_precise{};
        // namespace std {
        // template<> numeric_limits<super_precise> {
        //      double epsilon() {
        //            return 1e-30;
        //      }
        // }
        // }
        // float_equals<double>(0.0, 2e-30, 1); // returns true
        // float_equals<super_precise>(0.0, 2e-30, 1); // returns false
    }

    void test_float_equals_large()
    {
        const float compare_to = 20e6;
        // fp math with arguments of different magnitudes is wierd
        xlnt_assert(compare_to == compare_to + 1); // x == x + 1 ...
        xlnt_assert(compare_to != compare_to + 10); // x != x + 10
        xlnt_assert(compare_to != compare_to - 10); // x != x - 10
        // if the same epsilon was used for comparison of large values as the values around one
        // we'd have all the issues around zero again
        xlnt_assert(xlnt::detail::float_equals(compare_to, compare_to + 49));
        xlnt_assert(!xlnt::detail::float_equals(compare_to, compare_to + 50));
        xlnt_assert(xlnt::detail::float_equals(compare_to, compare_to - 49));
        xlnt_assert(!xlnt::detail::float_equals(compare_to, compare_to - 50));
        // float_equals also scales epsilon up to match the magnitude of its arguments
        // all the same options are available for increasing/decreasing the precision of the comparison
        // however the the epsilon source should always be of equal or lesser precision than the arguments when away from zero
    }

    void test_float_equals_nan()
    {
        const float nan = std::nanf("");
        // nans always compare false
        xlnt_assert(!xlnt::detail::float_equals(nan, 0.f));
        xlnt_assert(!xlnt::detail::float_equals(nan, nan));
        xlnt_assert(!xlnt::detail::float_equals(nan, 1000.f));
    }

    void test_float_equals_fairness()
    {
        // tests for parameter ordering dependency
        // (lhs ~= rhs) == (rhs ~= lhs)
        const double test_val = 1.0;
        const double test_diff_pass  = 1.192092e-07; // should all pass with this
        const double test_diff       = 1.192093e-07; // difference enough to provide different results if the comparison is not "fair"
        const double test_diff_fails = 1.192094e-07; // should all fail with this

        // test_diff_pass
        xlnt_assert(xlnt::detail::float_equals<float>((test_val + test_diff_pass), test_val, 1));
        xlnt_assert(xlnt::detail::float_equals<float>(test_val, (test_val + test_diff_pass), 1));
        xlnt_assert(xlnt::detail::float_equals<float>(-(test_val + test_diff_pass), -test_val, 1));
        xlnt_assert(xlnt::detail::float_equals<float>(-test_val, -(test_val + test_diff_pass), 1));
        // test_diff
        xlnt_assert(xlnt::detail::float_equals<float>((test_val + test_diff), test_val, 1));
        xlnt_assert(xlnt::detail::float_equals<float>(test_val, (test_val + test_diff), 1));
        xlnt_assert(xlnt::detail::float_equals<float>(-(test_val + test_diff), -test_val, 1));
        xlnt_assert(xlnt::detail::float_equals<float>(-test_val, -(test_val + test_diff), 1));
        // test_diff_fails
        xlnt_assert(!xlnt::detail::float_equals<float>((test_val + test_diff_fails), test_val, 1));
        xlnt_assert(!xlnt::detail::float_equals<float>(test_val, (test_val + test_diff_fails), 1));
        xlnt_assert(!xlnt::detail::float_equals<float>(-(test_val + test_diff_fails), -test_val, 1));
        xlnt_assert(!xlnt::detail::float_equals<float>(-test_val, -(test_val + test_diff_fails), 1));
    }

    void test_min()
    {
        // simple
        xlnt_assert(xlnt::detail::min(0, 1) == 0);
        xlnt_assert(xlnt::detail::min(1, 0) == 0);
        xlnt_assert(xlnt::detail::min(0.0, 1) == 0.0); // comparisons between different types just work
        xlnt_assert(xlnt::detail::min(1, 0.0) == 0.0);
        // negative numbers
        xlnt_assert(xlnt::detail::min(0, -1) == -1.0);
        xlnt_assert(xlnt::detail::min(-1, 0) == -1.0);
        xlnt_assert(xlnt::detail::min(0.0, -1) == -1.0);
        xlnt_assert(xlnt::detail::min(-1, 0.0) == -1.0);
        // no zeroes
        xlnt_assert(xlnt::detail::min(10, -10) == -10.0);
        xlnt_assert(xlnt::detail::min(-10, 10) == -10.0);
        xlnt_assert(xlnt::detail::min(10.0, -10) == -10.0);
        xlnt_assert(xlnt::detail::min(-10, 10.0) == -10.0);

        static_assert(xlnt::detail::min(-10, 10.0) == -10.0, "constexpr");
    }

    void test_max()
    {
        // simple
        xlnt_assert(xlnt::detail::max(0, 1) == 1);
        xlnt_assert(xlnt::detail::max(1, 0) == 1);
        xlnt_assert(xlnt::detail::max(0.0, 1) == 1.0); // comparisons between different types just work
        xlnt_assert(xlnt::detail::max(1, 0.0) == 1.0);
        // negative numbers
        xlnt_assert(xlnt::detail::max(0, -1) == 0.0);
        xlnt_assert(xlnt::detail::max(-1, 0) == 0.0);
        xlnt_assert(xlnt::detail::max(0.0, -1) == 0.0);
        xlnt_assert(xlnt::detail::max(-1, 0.0) == 0.0);
        // no zeroes
        xlnt_assert(xlnt::detail::max(10, -10) == 10.0);
        xlnt_assert(xlnt::detail::max(-10, 10) == 10.0);
        xlnt_assert(xlnt::detail::max(10.0, -10) == 10.0);
        xlnt_assert(xlnt::detail::max(-10, 10.0) == 10.0);

        static_assert(xlnt::detail::max(-10, 10.0) == 10.0, "constexpr");
    }

    void test_abs()
    {
        xlnt_assert(xlnt::detail::abs(0) == 0);
        xlnt_assert(xlnt::detail::abs(1) == 1);
        xlnt_assert(xlnt::detail::abs(-1) == 1);

        xlnt_assert(xlnt::detail::abs(0.0) == 0.0);
        xlnt_assert(xlnt::detail::abs(1.5) == 1.5);
        xlnt_assert(xlnt::detail::abs(-1.5) == 1.5);

        static_assert(xlnt::detail::abs(-1.23) == 1.23, "constexpr");
    }
};
static numeric_test_suite x;
