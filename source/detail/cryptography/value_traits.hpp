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

#pragma once

#include <string>

#include <detail/default_case.hpp>
#include <detail/cryptography/hash.hpp>
#include <detail/external/include_libstudxml.hpp>

namespace xml {

template <>
struct value_traits<xlnt::detail::hash_algorithm>
{
    static xlnt::detail::hash_algorithm parse(std::string hash_algorithm_string, const parser &)
    {
        if (hash_algorithm_string == "SHA1")
            return xlnt::detail::hash_algorithm::sha1;
        else if (hash_algorithm_string == "SHA256")
            return xlnt::detail::hash_algorithm::sha256;
        else if (hash_algorithm_string == "SHA384")
            return xlnt::detail::hash_algorithm::sha384;
        else if (hash_algorithm_string == "SHA512")
            return xlnt::detail::hash_algorithm::sha512;
        else if (hash_algorithm_string == "MD5")
            return xlnt::detail::hash_algorithm::md5;
        else if (hash_algorithm_string == "MD4")
            return xlnt::detail::hash_algorithm::md4;
        else if (hash_algorithm_string == "MD2")
            return xlnt::detail::hash_algorithm::md2;
        else if (hash_algorithm_string == "Ripemd128")
            return xlnt::detail::hash_algorithm::ripemd128;
        else if (hash_algorithm_string == "Ripemd160")
            return xlnt::detail::hash_algorithm::ripemd160;
        else if (hash_algorithm_string == "Whirlpool")
            return xlnt::detail::hash_algorithm::whirlpool;
        default_case(xlnt::detail::hash_algorithm::sha1);
    }

    static std::string serialize(xlnt::detail::hash_algorithm algorithm, const serializer &)
    {
        switch (algorithm)
        {
        case xlnt::detail::hash_algorithm::sha1:
            return "SHA1";
        case xlnt::detail::hash_algorithm::sha256:
            return "SHA256";
        case xlnt::detail::hash_algorithm::sha384:
            return "SHA384";
        case xlnt::detail::hash_algorithm::sha512:
            return "SHA512";
        case xlnt::detail::hash_algorithm::md5:
            return "MD5";
        case xlnt::detail::hash_algorithm::md4:
            return "MD4";
        case xlnt::detail::hash_algorithm::md2:
            return "MD2";
        case xlnt::detail::hash_algorithm::ripemd128:
            return "Ripemd128";
        case xlnt::detail::hash_algorithm::ripemd160:
            return "Ripemd160";
        case xlnt::detail::hash_algorithm::whirlpool:
            return "Whirlpool";
        }
        default_case("SHA1");
    }
}; // struct value_traits<xlnt::detail::hash_algorithm>

} // namespace xml
