// Copyright (c) 2014-2016 Thomas Fussell
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
#include <xlnt/utils/utf8string.hpp>

namespace xlnt {

utf8string utf8string::from_utf8(const std::string &s)
{
    utf8string result;
    std::copy(s.begin(), s.end(), std::back_inserter(result.bytes_));

    return result;
}

utf8string utf8string::from_latin1(const std::string &s)
{
    utf8string result;
    std::copy(s.begin(), s.end(), std::back_inserter(result.bytes_));

    return result;
}

utf8string utf8string::from_utf16(const std::string &s)
{
    utf8string result;
    utf8::utf16to8(s.begin(), s.end(), std::back_inserter(result.bytes_));

    return result;
}

utf8string utf8string::from_utf32(const std::string &s)
{
    utf8string result;
    utf8::utf32to8(s.begin(), s.end(), std::back_inserter(result.bytes_));

    return result;
}

bool utf8string::is_valid(const std::string &s)
{
    return utf8::is_valid(s.begin(), s.end());
}

bool utf8string::is_valid() const
{
    return utf8::is_valid(bytes_.begin(), bytes_.end());
}

} // namespace xlnt
