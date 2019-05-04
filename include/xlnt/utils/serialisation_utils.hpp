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

#if __cplusplus == 201703L || (defined(_MSC_VER) && _HAS_CXX17)

#include <charconv>

namespace xlnt {
/// <summary>
/// Takes in any nuber and outputs a string form of that number which will
/// serialise and deserialise without loss of precision
/// </summary>
template <typename Number>
std::string serialize_number_to_string(Number num)
{
    std::array<char, DBL_MAX_10_EXP + 1> str;
    auto result = std::to_chars(str.data(), str.data() + str.size(), num);
    *result.ptr = '\0';
    return str.data();
}
}

#else

namespace xlnt {
/// <summary>
/// Takes in any nuber and outputs a string form of that number which will
/// serialise and deserialise without loss of precision
/// </summary>
template <typename Number>
std::string serialize_number_to_string(Number num)
{
    static auto &facet = std::use_facet<std::num_put<char>>(std::locale::classic());
    // more digits and excel won't match
    constexpr int Excel_Digit_Precision = 15; //sf
    std::stringstream ss;
    ss.precision(Excel_Digit_Precision);
    facet.put(std::ostreambuf_iterator<char>(ss), ss, ss.fill(), num);
    return ss.str();
}
}

#endif
