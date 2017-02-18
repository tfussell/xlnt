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

#include <array>

#include <xlnt/utils/exceptions.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <detail/constants.hpp>
#include <detail/include_cryptopp.hpp>
#include <detail/include_libstudxml.hpp>
#include <detail/pole.hpp>
#include <detail/vector_streambuf.hpp>
#include <detail/xlsx_consumer.hpp>
#include <detail/xlsx_producer.hpp>

namespace xlnt {
namespace detail {

enum class hash_algorithm
{
    sha1,
    sha256,
    sha384,
    sha512,
    md5,
    md4,
    md2,
    ripemd128,
    ripemd160,
    whirlpool
};

struct crypto_helper
{
    static const std::size_t segment_length;

    enum class cipher_algorithm
    {
        aes,
        rc2,
        rc4,
        des,
        desx,
        triple_des,
        triple_des_112
    };

    enum class cipher_chaining
    {
        ecb, // electronic code book
        cbc, // cipher block chaining
        cfb // cipher feedback chaining
    };

    enum class cipher_direction
    {
        encryption,
        decryption
    };

    static std::vector<std::uint8_t> aes(const std::vector<std::uint8_t> &key, const std::vector<std::uint8_t> &iv,
        const std::vector<std::uint8_t> &source, cipher_chaining chaining, cipher_direction direction);

    static std::vector<std::uint8_t> decode_base64(const std::string &encoded);

    static std::string encode_base64(const std::vector<std::uint8_t> &decoded);

    static std::vector<std::uint8_t> hash(hash_algorithm algorithm, const std::vector<std::uint8_t> &input);

    static std::vector<std::uint8_t> file(POLE::Storage &storage, const std::string &name);

    struct standard_encryption_info
    {
        const std::size_t spin_count = 50000;
        std::size_t block_size;
        std::size_t key_bits;
        std::size_t key_bytes;
        std::size_t hash_size;
        cipher_algorithm cipher;
        cipher_chaining chaining;
        const hash_algorithm hash = hash_algorithm::sha1;
        std::vector<std::uint8_t> salt_value;
        std::vector<std::uint8_t> verifier_hash_input;
        std::vector<std::uint8_t> verifier_hash_value;
        std::vector<std::uint8_t> encrypted_key_value;
    };

    static std::vector<std::uint8_t> decrypt_xlsx_standard(const std::vector<std::uint8_t> &encryption_info,
        const std::string &password, const std::vector<std::uint8_t> &encrypted_package);

    struct agile_encryption_info
    {
        // key data
        struct
        {
            std::size_t salt_size;
            std::size_t block_size;
            std::size_t key_bits;
            std::size_t hash_size;
            std::string cipher_algorithm;
            std::string cipher_chaining;
            std::string hash_algorithm;
            std::vector<std::uint8_t> salt_value;
        } key_data;

        struct
        {
            std::vector<std::uint8_t> hmac_key;
            std::vector<std::uint8_t> hmac_value;
        } data_integrity;

        struct
        {
            std::size_t spin_count;
            std::size_t salt_size;
            std::size_t block_size;
            std::size_t key_bits;
            std::size_t hash_size;
            std::string cipher_algorithm;
            std::string cipher_chaining;
            hash_algorithm hash;
            std::vector<std::uint8_t> salt_value;
            std::vector<std::uint8_t> verifier_hash_input;
            std::vector<std::uint8_t> verifier_hash_value;
            std::vector<std::uint8_t> encrypted_key_value;
        } key_encryptor;
    };

    static agile_encryption_info generate_agile_encryption_info(const std::string &password);

    static std::vector<std::uint8_t> write_agile_encryption_info(const std::string &password);

    static std::vector<std::uint8_t> decrypt_xlsx_agile(const std::vector<std::uint8_t> &encryption_info,
        const std::string &password, const std::vector<std::uint8_t> &encrypted_package);

    static std::vector<std::uint8_t> decrypt_xlsx(const std::vector<std::uint8_t> &bytes, const std::string &password);

    static std::vector<std::uint8_t> encrypt_xlsx(const std::vector<std::uint8_t> &bytes, const std::string &password);
};

} // namespace detail
} // namespace xlnt
