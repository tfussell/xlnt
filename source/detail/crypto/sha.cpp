// Copyright (c) 2017 Thomas Fussell
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

#include <array>
#include <cstring>
#include <iomanip>
#include <string>
#include <sstream>

#include <detail/crypto/sha.hpp>

extern "C" {

extern void sha1_hash(const uint8_t *message, size_t len, uint32_t hash[5]);
extern void sha512_hash(const uint8_t *message, size_t len, uint64_t hash[8]);

}

namespace xlnt {
namespace detail {

std::vector<std::uint8_t> sha1(const std::vector<std::uint8_t> &data)
{
    std::array<std::uint32_t, 5> hash;
    sha1_hash(data.data(), data.size(), hash.data());

    std::vector<std::uint8_t> result(20, 0);
    auto result_iterator = result.begin();

    for (auto i : hash)
    {
        auto swapped = _byteswap_ulong(i);
        std::copy(
            reinterpret_cast<std::uint8_t *>(&swapped),
            reinterpret_cast<std::uint8_t *>(&swapped + 1),
            &*result_iterator);
        result_iterator += 4;
    }

    return result;
}

std::vector<std::uint8_t> sha512(const std::vector<std::uint8_t> &data)
{
    std::array<std::uint64_t, 8> hash;
    sha512_hash(data.data(), data.size(), hash.data());

    std::vector<std::uint8_t> result(64, 0);
    auto result_iterator = result.begin();

    for (auto i : hash)
    {
        auto swapped = _byteswap_uint64(i);
        std::copy(
            reinterpret_cast<std::uint8_t *>(&swapped),
            reinterpret_cast<std::uint8_t *>(&swapped + 1),
            &*result_iterator);
        result_iterator += 8;
    }

    return result;
}

} // namespace detail
} // namespace xlnt
