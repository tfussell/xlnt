// Copyright (c) 2016-2018 Thomas Fussell
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

#include <algorithm>
#include <iostream>
#include <vector>

#include <xlnt/xlnt_config.hpp>
#include <detail/cryptography/encryption_info.hpp>

namespace xlnt {
namespace detail {

/// <summary>
/// Allows a std::vector to be read through a std::istream.
/// </summary>
class XLNT_API crypt_istreambuf : public std::streambuf
{
    using int_type = std::streambuf::int_type;
    
public:
    crypt_istreambuf(std::istream &in_stream, const encryption_info &info);
    
    crypt_istreambuf(const crypt_istreambuf &) = delete;
    crypt_istreambuf &operator=(const crypt_istreambuf &) = delete;
    
private:
    int_type underflow();
    
    int_type uflow();
    
    std::streamsize showmanyc();
    
    std::streampos seekoff(std::streamoff off, std::ios_base::seekdir way, std::ios_base::openmode);
    
    std::streampos seekpos(std::streampos sp, std::ios_base::openmode);

private:
    void decrypt_block(std::size_t index);
    
    std::istream &in_stream_;
    std::size_t size_;
    std::size_t current_segment_;
    std::vector<std::uint8_t> buffer_;
    std::vector<std::uint8_t> key_;
    const encryption_info &info_;
};

/// <summary>
/// Allows a std::vector to be written through a std::ostream.
/// </summary>
class XLNT_API crypt_ostreambuf : public std::streambuf
{
    using int_type = std::streambuf::int_type;
    
public:
    crypt_ostreambuf(std::ostream &out_stream, const encryption_info &password);

    crypt_ostreambuf(const crypt_ostreambuf &) = delete;
    crypt_ostreambuf &operator=(const crypt_ostreambuf &) = delete;
    
private:
    int_type overflow(int_type c = traits_type::eof());
    
    std::streamsize xsputn(const char *s, std::streamsize n);
    
    std::streampos seekoff(std::streamoff off, std::ios_base::seekdir way, std::ios_base::openmode);
    
    std::streampos seekpos(std::streampos sp, std::ios_base::openmode);
    
private:
    std::ostream &out_stream_;
    std::vector<std::uint8_t> buffer_;
    std::vector<std::uint8_t> key_;
    const encryption_info &info_;
};

} // namespace detail
} // namespace xlnt
