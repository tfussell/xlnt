// Copyright (c) 2014-2018 Thomas Fussell
// Copyright (c) 2010-2015 openpyxl
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
#include <algorithm>
#include <locale>
#include <sstream>

#include <xlnt/worksheet/sheet_protection.hpp>

namespace {

template <typename T>
std::string int_to_hex(T i)
{
    std::stringstream stream;
    stream << std::hex << i;

    return stream.str();
}

} // namespace

namespace xlnt {

void sheet_protection::password(const std::string &password)
{
    hashed_password_ = hash_password(password);
}

std::string sheet_protection::hashed_password() const
{
    return hashed_password_;
}

std::string sheet_protection::hash_password(const std::string &plaintext_password)
{
    int password = 0x0000;
    int i = 1;

    for (auto character : plaintext_password)
    {
        int value = character << i;
        int rotated_bits = value >> 15;
        value &= 0x7fff;
        password ^= (value | rotated_bits);
        i++;
    }

    password ^= plaintext_password.size();
    password ^= 0xCE4B;

    std::string hashed = int_to_hex(password);
    std::transform(hashed.begin(), hashed.end(), hashed.begin(),
        [](char c) { return std::toupper(c, std::locale::classic()); });

    return hashed;
}
}
