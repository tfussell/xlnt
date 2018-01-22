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

namespace xlnt {
namespace detail {

/// <summary>
/// Allows a std::vector to be read through a std::istream.
/// </summary>
class XLNT_API vector_istreambuf : public std::streambuf
{
    using int_type = std::streambuf::int_type;

public:
    vector_istreambuf(const std::vector<std::uint8_t> &data);

    vector_istreambuf(const vector_istreambuf &) = delete;
    vector_istreambuf &operator=(const vector_istreambuf &) = delete;

private:
    int_type underflow();

    int_type uflow();

    std::streamsize showmanyc();

    std::streampos seekoff(std::streamoff off, std::ios_base::seekdir way, std::ios_base::openmode);

    std::streampos seekpos(std::streampos sp, std::ios_base::openmode);

private:
    const std::vector<std::uint8_t> &data_;
    std::size_t position_;
};

/// <summary>
/// Allows a std::vector to be written through a std::ostream.
/// </summary>
class XLNT_API vector_ostreambuf : public std::streambuf
{
    using int_type = std::streambuf::int_type;

public:
    vector_ostreambuf(std::vector<std::uint8_t> &data);

    vector_ostreambuf(const vector_ostreambuf &) = delete;
    vector_ostreambuf &operator=(const vector_ostreambuf &) = delete;

private:
    int_type overflow(int_type c = traits_type::eof());

    std::streamsize xsputn(const char *s, std::streamsize n);

    std::streampos seekoff(std::streamoff off, std::ios_base::seekdir way, std::ios_base::openmode);

    std::streampos seekpos(std::streampos sp, std::ios_base::openmode);

private:
    std::vector<std::uint8_t> &data_;
    std::size_t position_;
};

//TODO: detail headers shouldn't be exporting such functions

/// <summary>
/// Helper function to read all data from in_stream and store them in a vector.
/// </summary>
XLNT_API std::vector<std::uint8_t> to_vector(std::istream &in_stream);

/// <summary>
/// Helper function to write all data from bytes into out_stream.
/// </summary>
XLNT_API void to_stream(const std::vector<std::uint8_t> &bytes, std::ostream &out_stream);

/// <summary>
/// Shortcut function to stream a vector of bytes into a std::ostream.
/// </summary>
XLNT_API std::ostream &operator<<(std::ostream &out_stream, const std::vector<std::uint8_t> &bytes);

} // namespace detail
} // namespace xlnt
