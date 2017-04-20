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

#include <detail/cryptography/hash.hpp>
#include <detail/cryptography/sha.hpp>

namespace xlnt {
namespace detail {

void hash(hash_algorithm algorithm, const std::vector<std::uint8_t> &input, std::vector<std::uint8_t> &output)
{
    if (algorithm == hash_algorithm::sha512)
    {
        xlnt::detail::sha512(input, output);
    }
    else if (algorithm == hash_algorithm::sha1)
    {
        xlnt::detail::sha1(input, output);
    }
    else
    {
        throw xlnt::exception("unsupported hash algorithm");
    }
}

std::vector<std::uint8_t> hash(hash_algorithm algorithm, const std::vector<std::uint8_t> &input)
{
    auto output = std::vector<std::uint8_t>();
    hash(algorithm, input, output);

    return output;
}

}; // namespace detail
}; // namespace xlnt
