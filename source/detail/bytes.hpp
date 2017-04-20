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

#include <cstdint>
#include <vector>

namespace xlnt {
namespace detail {

using byte = std::uint8_t;
using byte_vector = std::vector<byte>;

template <typename T>
auto read_int(const byte_vector &raw_data, std::size_t &index)
{
    auto result = *reinterpret_cast<const T *>(&raw_data[index]);
    index += sizeof(T);

    return result;
}

template <typename T>
byte *vector_byte(std::vector<T> &v, std::size_t offset)
{
    return reinterpret_cast<byte *>(v.data() + offset);
}

template <typename T>
byte *first_byte(std::vector<T> &v)
{
    return vector_byte(v, 0);
}

template <typename T>
byte *last_byte(std::vector<T> &v)
{
    return vector_byte(v, v.size());
}

} // namespace detail
} // namespace xlnt
