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

#include <array>
#include <cmath>
#include <cstdint>
#include <vector>

#include <detail/binary.hpp>
#include <detail/constants.hpp>
#include <detail/unicode.hpp>
#include <detail/cryptography/aes.hpp>
#include <detail/cryptography/base64.hpp>
#include <detail/cryptography/compound_document.hpp>
#include <detail/cryptography/crypt_stream.hpp>
#include <detail/cryptography/encryption_info.hpp>
#include <detail/cryptography/hash.hpp>
#include <detail/cryptography/value_traits.hpp>
#include <detail/cryptography/xlsx_crypto_consumer.hpp>
#include <detail/external/include_libstudxml.hpp>
#include <detail/serialization/vector_streambuf.hpp>
#include <detail/serialization/xlsx_consumer.hpp>
#include <xlnt/utils/exceptions.hpp>

namespace {

using xlnt::detail::byte;
using xlnt::detail::read;
using xlnt::detail::encryption_info;

} // namespace

namespace xlnt {
namespace detail {

crypt_istreambuf::crypt_istreambuf(std::istream &in_stream,
    const encryption_info &info)
    : in_stream_(in_stream),
      size_(read<std::uint64_t>(in_stream_)),
      key_(info.calculate_key()),
      current_segment_(std::numeric_limits<std::size_t>::max()),
      info_(info)
{
}

crypt_istreambuf::int_type crypt_istreambuf::underflow()
{
    return in_stream_.peek();
}

crypt_istreambuf::int_type crypt_istreambuf::uflow()
{
    auto index = in_stream_.tellg();
    if (index > static_cast<std::ios::streamoff>(size_))
    {
        index = size_;
    }
    auto segment = index / 4096;
    
    if (static_cast<std::size_t>(segment) != current_segment_)
    {
        decrypt_block(segment);
        current_segment_ = segment;
    }
    
    in_stream_.seekg(index + static_cast<std::streamoff>(1));

    return buffer_.at(index % 4096);
}

std::streamsize crypt_istreambuf::showmanyc()
{
    return 0;
}

std::streampos crypt_istreambuf::seekoff(std::streamoff off, std::ios_base::seekdir way, std::ios_base::openmode)
{
    in_stream_.seekg(off, way);

    auto index = in_stream_.tellg();
    if (index > static_cast<std::ios::streamoff>(size_))
    {
        index = size_;
    }
    auto segment = index / 4096;
    
    if (static_cast<std::size_t>(segment) != current_segment_)
    {
        decrypt_block(segment);
        current_segment_ = segment;
    }

    return index;
}

std::streampos crypt_istreambuf::seekpos(std::streampos sp, std::ios_base::openmode)
{
    in_stream_.seekg(sp);
    return in_stream_.tellg();
}
    
void crypt_istreambuf::decrypt_block(std::size_t index)
{
    const std::size_t segment_length = 4096;
    const auto start = index * segment_length;
    const auto end = std::min((index + 1) * segment_length, size_);

    buffer_.resize(segment_length, 0);
    in_stream_.seekg(start, std::ios::beg);

    if (info_.is_agile)
    {
        auto salt_with_block_key = info_.agile.key_encryptor.salt_value;
        salt_with_block_key.resize(info_.agile.key_encryptor.salt_size + sizeof(std::uint32_t), 0);
        
        auto &block_key_index = *reinterpret_cast<std::uint32_t *>(salt_with_block_key.data()
            + info_.agile.key_encryptor.salt_size);
        block_key_index = static_cast<std::uint32_t>(index);

        auto iv = hash(info_.agile.key_encryptor.hash, salt_with_block_key);
        iv.resize(16);

        in_stream_.read(reinterpret_cast<char *>(buffer_.data()), segment_length);
        buffer_ = xlnt::detail::aes_cbc_decrypt(buffer_, key_, iv);
    }
    else
    {
        in_stream_.read(reinterpret_cast<char *>(buffer_.data()),
            static_cast<std::streamsize>(buffer_.size()));

        buffer_ = xlnt::detail::aes_ecb_decrypt(buffer_, key_);
    }

    const auto size = (end > start) ? (end - start) : 0;
    buffer_.resize(size);
}

crypt_ostreambuf::crypt_ostreambuf(std::ostream &out_stream,
    const encryption_info &info)
    : out_stream_(out_stream),
      key_(info.calculate_key()),
      info_(info)
{
}

crypt_ostreambuf::int_type crypt_ostreambuf::overflow(int_type c)
{
    out_stream_.put(c);
    return static_cast<crypt_ostreambuf::int_type>(out_stream_.tellp());
}

std::streamsize crypt_ostreambuf::xsputn(const char *s, std::streamsize n)
{
    out_stream_.write(s, n);
    return out_stream_.tellp();
}

std::streampos crypt_ostreambuf::seekoff(std::streamoff off, std::ios_base::seekdir way, std::ios_base::openmode)
{
    out_stream_.seekp(off, way);
    return out_stream_.tellp();
}

std::streampos crypt_ostreambuf::seekpos(std::streampos sp, std::ios_base::openmode)
{
    out_stream_.seekp(sp);
    return out_stream_.tellp();
}

} // namespace detail
} // namespace xlnt
