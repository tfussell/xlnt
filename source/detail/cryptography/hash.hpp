// Copyright (c) 2017-2018 Thomas Fussell
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

#include <vector>

#include <detail/cryptography/sha.hpp>
#include <xlnt/utils/exceptions.hpp>

namespace xlnt {
namespace detail {

enum class hash_algorithm
{
    sha1,
    sha256,
    sha384,
    sha512,
    md2,
    md4,
    md5,
    ripemd128,
    ripemd160,
    whirlpool
};

void hash(hash_algorithm algorithm, const std::vector<std::uint8_t> &input, std::vector<std::uint8_t> &output);
std::vector<std::uint8_t> hash(hash_algorithm algorithm, const std::vector<std::uint8_t> &input);

}; // namespace detail
}; // namespace xlnt

