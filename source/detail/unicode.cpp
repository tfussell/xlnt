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

#pragma clang diagnostic push
#pragma clang diagnostic ignored "-Wdeprecated"
#pragma clang diagnostic ignored "-Wweak-vtables"
#pragma clang diagnostic ignored "-Wsign-conversion"
#include <utf8.h>
#pragma clang diagnostic pop

#include <detail/unicode.hpp>

namespace xlnt {
namespace detail {

std::u16string utf8_to_utf16(const std::string &utf8_string)
{
    std::u16string result;
    utf8::utf8to16(utf8_string.begin(), utf8_string.end(), std::back_inserter(result));

    return result;
}

std::string utf16_to_utf8(const std::u16string &utf16_string)
{
    std::string result;
    utf8::utf16to8(utf16_string.begin(), utf16_string.end(), std::back_inserter(result));

    return result;
}

std::string latin1_to_utf8(const std::string &latin1)
{
    std::string utf8;

    for (auto character : latin1)
    {
        if (character >= 0)
        {
            utf8.push_back(character);
        }
        else
        {
            utf8.push_back(static_cast<char>(0xc0 + (character >> 6)));
            utf8.push_back(static_cast<char>(0x80 + (character & 0x3f)));
        }
    }

    return utf8;
}

} // namespace detail
} // namespace xlnt
