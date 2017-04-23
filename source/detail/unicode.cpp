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

#include <codecvt>
#include <locale>
#include <string>

#include <detail/unicode.hpp>

namespace xlnt {
namespace detail {

#ifdef _MSC_VER
std::u16string utf8_to_utf16(const std::string &utf8_string)
{
    // use wchar_t instead of char16_t on Windows because of a bug in MSVC
    // error LNK2001: unresolved external symbol std::codecvt::id
    // https://connect.microsoft.com/VisualStudio/Feedback/Details/1403302
    auto converted = std::wstring_convert<std::codecvt_utf8_utf16<wchar_t>,
        wchar_t>{}.from_bytes(utf8_string);
    return std::u16string(converted.begin(), converted.end());
}

std::string utf16_to_utf8(const std::u16string &utf16_string)
{
    std::wstring utf16_wstring(utf16_string.begin(), utf16_string.end());
    // use wchar_t instead of char16_t on Windows because of a bug in MSVC
    // error LNK2001: unresolved external symbol std::codecvt::id
    // https://connect.microsoft.com/VisualStudio/Feedback/Details/1403302
    return std::wstring_convert<std::codecvt_utf8_utf16<wchar_t>,
        wchar_t>{}.to_bytes(utf16_wstring);
}
#else
std::u16string utf8_to_utf16(const std::string &utf8_string)
{
    return std::wstring_convert<std::codecvt_utf8_utf16<char16_t>,
        char16_t>{}.from_bytes(utf8_string);
}

std::string utf16_to_utf8(const std::u16string &utf16_string)
{
    return std::wstring_convert<std::codecvt_utf8_utf16<char16_t>,
        char16_t>{}.to_bytes(utf16_string);
}
#endif

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
