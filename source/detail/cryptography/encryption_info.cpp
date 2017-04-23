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

#include <detail/binary.hpp>
#include <detail/cryptography/aes.hpp>
#include <detail/cryptography/encryption_info.hpp>

namespace {

using xlnt::detail::encryption_info;

std::vector<std::uint8_t> calculate_standard_key(
    encryption_info::standard_encryption_info info,
    const std::u16string &password)
{
    // H_0 = H(salt + password)
    auto salt_plus_password = info.salt;
    auto password_bytes = xlnt::detail::string_to_bytes(password);
    std::copy(password_bytes.begin(), 
        password_bytes.end(), 
        std::back_inserter(salt_plus_password));
    auto h_0 = hash(info.hash, salt_plus_password);

    // H_n = H(iterator + H_n-1)
    std::vector<std::uint8_t> iterator_plus_h_n(4, 0);
    iterator_plus_h_n.insert(iterator_plus_h_n.end(), h_0.begin(), h_0.end());
    std::uint32_t &iterator = *reinterpret_cast<std::uint32_t *>(iterator_plus_h_n.data());
    std::vector<std::uint8_t> h_n;
    for (iterator = 0; iterator < info.spin_count; ++iterator)
    {
        hash(info.hash, iterator_plus_h_n, h_n);
        std::copy(h_n.begin(), h_n.end(), iterator_plus_h_n.begin() + 4);
    }

    // H_final = H(H_n + block)
    auto h_n_plus_block = h_n;
    const std::uint32_t block_number = 0;
    h_n_plus_block.insert(
        h_n_plus_block.end(),
        reinterpret_cast<const std::uint8_t *>(&block_number),
        reinterpret_cast<const std::uint8_t *>(&block_number) + sizeof(std::uint32_t));
    auto h_final = hash(info.hash, h_n_plus_block);

    // X1 = H(h_final ^ 0x36)
    std::vector<std::uint8_t> buffer(64, 0x36);
    for (std::size_t i = 0; i < h_final.size(); ++i)
    {
        buffer[i] = static_cast<std::uint8_t>(0x36 ^ h_final[i]);
    }
    auto X1 = hash(info.hash, buffer);

    // X2 = H(h_final ^ 0x5C)
    buffer.assign(64, 0x5c);
    for (std::size_t i = 0; i < h_final.size(); ++i)
    {
        buffer[i] = static_cast<std::uint8_t>(0x5c ^ h_final[i]);
    }
    auto X2 = hash(info.hash, buffer);

    auto X3 = X1;
    X3.insert(X3.end(), X2.begin(), X2.end());

    auto key = std::vector<std::uint8_t>(X3.begin(),
        X3.begin() + static_cast<std::ptrdiff_t>(info.key_bytes));

    using xlnt::detail::aes_ecb_decrypt;

    auto calculated_verifier_hash = hash(info.hash,
        aes_ecb_decrypt(info.encrypted_verifier, key));
    auto decrypted_verifier_hash = aes_ecb_decrypt(
        info.encrypted_verifier_hash, key);
    decrypted_verifier_hash.resize(calculated_verifier_hash.size());

    if (calculated_verifier_hash != decrypted_verifier_hash)
    {
        throw xlnt::exception("bad password");
    }

    return key;
}

std::vector<std::uint8_t> calculate_agile_key(
    encryption_info::agile_encryption_info info,
    const std::u16string &password)
{
    // H_0 = H(salt + password)
    auto salt_plus_password = info.key_encryptor.salt_value;
    auto password_bytes = xlnt::detail::string_to_bytes(password);
    std::copy(password_bytes.begin(),
        password_bytes.end(),
        std::back_inserter(salt_plus_password));

    auto h_0 = hash(info.key_encryptor.hash, salt_plus_password);

    // H_n = H(iterator + H_n-1)
    std::vector<std::uint8_t> iterator_plus_h_n(4, 0);
    iterator_plus_h_n.insert(iterator_plus_h_n.end(), h_0.begin(), h_0.end());
    std::uint32_t &iterator = *reinterpret_cast<std::uint32_t *>(iterator_plus_h_n.data());
    std::vector<std::uint8_t> h_n;
    for (iterator = 0; iterator < info.key_encryptor.spin_count; ++iterator)
    {
        hash(info.key_encryptor.hash, iterator_plus_h_n, h_n);
        std::copy(h_n.begin(), h_n.end(), iterator_plus_h_n.begin() + 4);
    }

    static const std::size_t block_size = 8;

    auto calculate_block = [&info](
        const std::vector<std::uint8_t> &raw_key,
        const std::array<std::uint8_t, block_size> &block,
        const std::vector<std::uint8_t> &encrypted)
    {
        auto combined = raw_key;
        combined.insert(combined.end(), block.begin(), block.end());

        auto key = hash(info.key_encryptor.hash, combined);
        key.resize(info.key_encryptor.key_bits / 8);

        using xlnt::detail::aes_cbc_decrypt;

        return aes_cbc_decrypt(encrypted, key, info.key_encryptor.salt_value);
    };

    const std::array<std::uint8_t, block_size> input_block_key = { { 0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79 } };
    auto hash_input = calculate_block(h_n, input_block_key, info.key_encryptor.verifier_hash_input);
    auto calculated_verifier = hash(info.key_encryptor.hash, hash_input);

    const std::array<std::uint8_t, block_size> verifier_block_key = {
        { 0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e } };
    auto expected_verifier = calculate_block(h_n, verifier_block_key, info.key_encryptor.verifier_hash_value);
    expected_verifier.resize(calculated_verifier.size());

    if (calculated_verifier != expected_verifier)
    {
        throw xlnt::exception("bad password");
    }

    const std::array<std::uint8_t, block_size> key_value_block_key =
    {
        { 0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6 }
    };

    return calculate_block(h_n, key_value_block_key, info.key_encryptor.encrypted_key_value);
}

} // namespace

namespace xlnt {
namespace detail {

std::vector<std::uint8_t> encryption_info::calculate_key() const
{
    return is_agile
        ? calculate_agile_key(agile, password)
        : calculate_standard_key(standard, password);
}

} // namespace detail
} // namespace xlnt
