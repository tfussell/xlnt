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

#pragma once

#include <xlnt/xlnt_config.hpp>
#include <sstream>
#include <type_traits>

namespace xlnt {
namespace detail {
/// <summary>
/// Takes in any number and outputs a string form of that number which will
/// serialise and deserialise without loss of precision
/// </summary>
template <typename Number>
std::string serialize_number_to_string(Number num)
{
    // more digits and excel won't match
    constexpr int Excel_Digit_Precision = 15; //sf
    std::stringstream ss;
    ss.precision(Excel_Digit_Precision);
    ss << num;
    return ss.str();
}

/// <summary>
/// constexpr abs
/// </summary>
template <typename Number>
constexpr Number abs(Number val)
{
    if (val < Number{0})
    {
        return -val;
    }
    return val;
};

/// <summary>
/// constexpr max
/// </summary>
template <typename Number>
constexpr Number max(Number lval, Number rval)
{
    if (lval < rval)
    {
        return rval;
    }
    return lval;
};

/// <summary>
/// Floating point equality requires a bit of fuzzingdue to the imprecise nature of fp calculation
/// </summary>
template <typename EpsilonType = float, // the type to extract epsilon from
    typename LNumber, typename RNumber> // parameter types (deduced)
constexpr bool
float_equals(const LNumber &lhs, const RNumber &rhs,
    int epsilon_scale = 100) // scale the "fuzzy" equality. Higher value gives a more tolerant comparison
{
    static_assert(std::is_floating_point<LNumber>::value || std::is_floating_point<RNumber>::value,
        "Using this function with two integers is just wasting time. Use ==");
    static_assert(std::is_floating_point<EpsilonType>::value,
        "Cannot extract epsilon from a number that isn't a floating point type");

    // NANs always compare false with themselves
    if ((lhs != lhs) || (rhs != rhs)) // std::isnan isn't constexpr
    {
        return false;
    }
    // a type that lhs and rhs can agree on
    using common_t = std::common_type<LNumber, RNumber>::type;
    // epsilon type defaults to float because even if both args are a higher precision type
    // either or both could have been promoted by prior operations
    // if a higher precision is required, the template type can be changed
    constexpr common_t epsilon = std::numeric_limits<EpsilonType>::epsilon();
    // the "epsilon" then needs to be scaled into the comparison range
    // epsilon for numeric_limits is valid when abs(x) <1.0, scaling only needs to be upwards
    // in particular, this prevents a lhs of 0 from requiring an exact comparison
    // additionally, a scale factor is applied.
    common_t scaled_fuzz = epsilon_scale * epsilon * max(xlnt::detail::abs<common_t>(lhs), common_t{1});
    return ((lhs + scaled_fuzz) >= rhs) && ((rhs + scaled_fuzz) >= lhs);
}

static_assert(0.1 != 0.1f, "Built in equality fails when comparing float and double due to imprecision");
static_assert(float_equals(0.1, 0.1f), "fuzzy equality allows comparison between double and float");
} // namespace detail
} // namespace xlnt
