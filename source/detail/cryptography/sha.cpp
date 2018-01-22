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

#include <algorithm>
#include <array>
#include <cstring>
#include <iomanip>
#include <string>
#include <sstream>

#include <detail/cryptography/sha.hpp>

extern "C" {

extern void sha1_hash(const uint8_t *message, size_t len, uint32_t hash[5]);
extern void sha512_hash(const uint8_t *message, size_t len, uint64_t hash[8]);

}

namespace {

#ifdef _MSC_VER
inline std::uint32_t byteswap32(std::uint32_t i)
{
    return _byteswap_ulong(i);
}

inline std::uint64_t byteswap64(std::uint64_t i)
{
    return _byteswap_uint64(i);
}
#else
inline std::uint32_t byteswap32(std::uint32_t i)
{
    return __builtin_bswap32(i);
}

inline std::uint64_t byteswap64(std::uint64_t i)
{
    return __builtin_bswap64(i);
}
#endif

void byteswap(std::uint32_t *arr, std::size_t len)
{
    for (auto i = std::size_t(0); i < len; ++i)
    {
        arr[i] = byteswap32(arr[i]);
    }
}

void byteswap(std::uint64_t *arr, std::size_t len)
{
    for (auto i = std::size_t(0); i < len; ++i)
    {
        arr[i] = byteswap64(arr[i]);
    }
}

} // namespace

namespace xlnt {
namespace detail {

void sha1(const std::vector<std::uint8_t> &input, std::vector<std::uint8_t> &output)
{
    static const auto sha1_bytes = 20;

    output.resize(sha1_bytes);
    auto output_pointer_u32 = reinterpret_cast<std::uint32_t *>(output.data());

    sha1_hash(input.data(), input.size(), output_pointer_u32);

    byteswap(output_pointer_u32, sha1_bytes / sizeof(std::uint32_t));
}

void sha512(const std::vector<std::uint8_t> &input, std::vector<std::uint8_t> &output)
{
    static const auto sha512_bytes = 64;

    output.resize(sha512_bytes);
    auto output_pointer_u64 = reinterpret_cast<std::uint64_t *>(output.data());

    sha512_hash(input.data(), input.size(), output_pointer_u64);

    byteswap(output_pointer_u64, sha512_bytes / sizeof(std::uint64_t));
}

} // namespace detail
} // namespace xlnt
