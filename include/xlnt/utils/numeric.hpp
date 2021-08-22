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

#pragma once

#include <xlnt/xlnt_config.hpp>
#include <algorithm>
#include <cassert>
#include <cmath>
#include <cstddef>
#include <limits>
#include <sstream>
#include <type_traits>

#undef min
#undef max

namespace xlnt {
namespace detail {

/// <summary>
/// constexpr abs
/// </summary>
template <typename Number>
constexpr Number abs(Number val)
{
    return (val < Number{0}) ? -val : val;
}

/// <summary>
/// constexpr max
/// </summary>
template <typename NumberL, typename NumberR>
constexpr typename std::common_type<NumberL, NumberR>::type max(NumberL lval, NumberR rval)
{
    return (lval < rval) ? rval : lval;
}

/// <summary>
/// constexpr min
/// </summary>
template <typename NumberL, typename NumberR>
constexpr typename std::common_type<NumberL, NumberR>::type min(NumberL lval, NumberR rval)
{
    return (lval < rval) ? lval : rval;
}

/// <summary>
/// Floating point equality requires a bit of fuzzing due to the imprecise nature of fp calculation
/// References:
/// - Several blogs/articles were referenced with the following being the most useful
/// -- https://randomascii.wordpress.com/2012/02/25/comparing-floating-point-numbers-2012-edition/
/// -- http://realtimecollisiondetection.net/blog/?p=89
/// - Testing Frameworks {Catch2, Boost, Google}, primarily for selecting the default scale factor
/// -- None of these even remotely agree
/// </summary>
template <typename EpsilonType = float, // the type to extract epsilon from
    typename LNumber, typename RNumber> // parameter types (deduced)
bool float_equals(const LNumber &lhs, const RNumber &rhs,
    int epsilon_scale = 20) // scale the "fuzzy" equality. Higher value gives a more tolerant comparison
{
    // a type that lhs and rhs can agree on
    using common_t = typename std::common_type<LNumber, RNumber>::type;
    // asserts for sane usage
    static_assert(std::is_floating_point<LNumber>::value || std::is_floating_point<RNumber>::value,
        "Using this function with two integers is just wasting time. Use ==");
    static_assert(std::numeric_limits<EpsilonType>::epsilon() < EpsilonType{1},
        "epsilon >= 1.0 will cause all comparisons to return true");

    // NANs always compare false with themselves
    if (std::isnan(lhs) || std::isnan(rhs))
    {
        return false;
    }
    // epsilon type defaults to float because even if both args are a higher precision type
    // either or both could have been promoted by prior operations
    // if a higher precision is required, the template type can be changed
    constexpr common_t epsilon = static_cast<common_t>(std::numeric_limits<EpsilonType>::epsilon());
    // the "epsilon" then needs to be scaled into the comparison range
    // epsilon for numeric_limits is valid when abs(x) <1.0, scaling only needs to be upwards
    // in particular, this prevents a lhs of 0 from requiring an exact comparison
    // additionally, a scale factor is applied.
    common_t scaled_fuzz = epsilon_scale * epsilon * max(max(xlnt::detail::abs<common_t>(lhs),
                                                             xlnt::detail::abs<common_t>(rhs)), // |max| of parameters.
                               common_t{1}); // clamp
    return ((lhs + scaled_fuzz) >= rhs) && ((rhs + scaled_fuzz) >= lhs);
}

class number_serialiser
{
    static constexpr int Excel_Digit_Precision = 15; //sf
    bool should_convert_comma;

    static void convert_comma_to_pt(char *buf, int len)
    {
        char *buf_end = buf + len;
        char *decimal = std::find(buf, buf_end, ',');
        if (decimal != buf_end)
        {
            *decimal = '.';
        }
    }

    static void convert_pt_to_comma(char *buf, size_t len)
    {
        char *buf_end = buf + len;
        char *decimal = std::find(buf, buf_end, '.');
        if (decimal != buf_end)
        {
            *decimal = ',';
        }
    }

public:
    explicit number_serialiser()
        : should_convert_comma(localeconv()->decimal_point[0] == ',')
    {
    }

    // for printing to file.
    // This matches the output format of excel irrespective of current locale
    std::string serialise(double d) const
    {
        char buf[30];
        int len = snprintf(buf, sizeof(buf), "%.15g", d);
        if (should_convert_comma)
        {
            convert_comma_to_pt(buf, len);
        }
        return std::string(buf, static_cast<size_t>(len));
    }

    // replacement for std::to_string / s*printf("%f", ...)
    // behaves same irrespective of locale
    std::string serialise_short(double d) const
    {
        char buf[30];
        int len = snprintf(buf, sizeof(buf), "%f", d);
        if (should_convert_comma)
        {
            convert_comma_to_pt(buf, len);
        }
        return std::string(buf, static_cast<size_t>(len));
    }

    double deserialise(const std::string &s, ptrdiff_t *len_converted) const
    {
        assert(!s.empty());
        assert(len_converted != nullptr);
        char *end_of_convert;
        if (!should_convert_comma)
        {
            double d = strtod(s.c_str(), &end_of_convert);
            *len_converted = end_of_convert - s.c_str();
            return d;
        }
        char buf[30];
        assert(s.size() < sizeof(buf));
        auto copy_end = std::copy(s.begin(), s.end(), buf);
        convert_pt_to_comma(buf, static_cast<size_t>(copy_end - buf));
        double d = strtod(buf, &end_of_convert);
        *len_converted = end_of_convert - buf;
        return d;
    }

    double deserialise(const std::string &s) const
    {
        ptrdiff_t ignore;
        return deserialise(s, &ignore);
    }
};

} // namespace detail
} // namespace xlnt
