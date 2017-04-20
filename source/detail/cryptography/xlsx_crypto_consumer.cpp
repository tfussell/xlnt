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
#include <cstdint>
#include <vector>

#include <detail/bytes.hpp>
#include <detail/constants.hpp>
#include <detail/unicode.hpp>
#include <detail/cryptography/encryption_info.hpp>
#include <detail/cryptography/aes.hpp>
#include <detail/cryptography/base64.hpp>
#include <detail/cryptography/pole.hpp>
#include <detail/cryptography/value_traits.hpp>
#include <detail/external/include_libstudxml.hpp>
#include <detail/serialization/vector_streambuf.hpp>
#include <detail/serialization/xlsx_consumer.hpp>
#include <xlnt/utils/exceptions.hpp>

namespace {

using xlnt::detail::byte;
using xlnt::detail::byte_vector;
using xlnt::detail::encryption_info;

std::vector<std::uint8_t> decrypt_xlsx_standard(
    const byte_vector &encryption_info,
    const std::u16string &password,
    const byte_vector &encrypted_package)
{
    std::size_t offset = 0;

    encryption_info::standard_encryption_info info;

    using xlnt::detail::read_int;

    auto header_length = read_int<std::uint32_t>(encryption_info, offset);
    auto index_at_start = offset;
    /*auto skip_flags = */ read_int<std::uint32_t>(encryption_info, offset);
    /*auto size_extra = */ read_int<std::uint32_t>(encryption_info, offset);
    auto alg_id = read_int<std::uint32_t>(encryption_info, offset);

    if (alg_id == 0 || alg_id == 0x0000660E || alg_id == 0x0000660F || alg_id == 0x00006610)
    {
        info.cipher = xlnt::detail::cipher_algorithm::aes;
    }
    else
    {
        throw xlnt::exception("invalid cipher algorithm");
    }

    auto alg_id_hash = read_int<std::uint32_t>(encryption_info, offset);
    if (alg_id_hash != 0x00008004 && alg_id_hash == 0)
    {
        throw xlnt::exception("invalid hash algorithm");
    }

    info.key_bits = read_int<std::uint32_t>(encryption_info, offset);
    info.key_bytes = info.key_bits / 8;

    auto provider_type = read_int<std::uint32_t>(encryption_info, offset);
    if (provider_type != 0 && provider_type != 0x00000018)
    {
        throw xlnt::exception("invalid provider type");
    }

    read_int<std::uint32_t>(encryption_info, offset); // reserved 1
    if (read_int<std::uint32_t>(encryption_info, offset) != 0) // reserved 2
    {
        throw xlnt::exception("invalid header");
    }

    const auto csp_name_length = header_length - (offset - index_at_start);
    std::vector<std::uint16_t> csp_name_wide(
        reinterpret_cast<const std::uint16_t *>(&*(encryption_info.begin() + static_cast<std::ptrdiff_t>(offset))),
        reinterpret_cast<const std::uint16_t *>(
            &*(encryption_info.begin() + static_cast<std::ptrdiff_t>(offset + csp_name_length))));
    std::string csp_name(csp_name_wide.begin(), csp_name_wide.end() - 1); // without trailing null
    if (csp_name != "Microsoft Enhanced RSA and AES Cryptographic Provider (Prototype)"
        && csp_name != "Microsoft Enhanced RSA and AES Cryptographic Provider")
    {
        throw xlnt::exception("invalid cryptographic provider");
    }
    offset += csp_name_length;

    const auto salt_size = read_int<std::uint32_t>(encryption_info, offset);
    std::vector<std::uint8_t> salt(encryption_info.begin() + static_cast<std::ptrdiff_t>(offset),
        encryption_info.begin() + static_cast<std::ptrdiff_t>(offset + salt_size));
    offset += salt_size;

    static const auto verifier_size = std::size_t(16);
    std::vector<std::uint8_t> encrypted_verifier(encryption_info.begin() + static_cast<std::ptrdiff_t>(offset),
        encryption_info.begin() + static_cast<std::ptrdiff_t>(offset + verifier_size));
    offset += verifier_size;

    const auto verifier_hash_size = read_int<std::uint32_t>(encryption_info, offset);
    const auto encrypted_verifier_hash_size = std::size_t(32);
    std::vector<std::uint8_t> encrypted_verifier_hash(encryption_info.begin() + static_cast<std::ptrdiff_t>(offset),
        encryption_info.begin() + static_cast<std::ptrdiff_t>(offset + encrypted_verifier_hash_size));
    offset += encrypted_verifier_hash_size;

    if (offset != encryption_info.size())
    {
        throw xlnt::exception("extra data after encryption info");
    }

    // begin key generation algorithm

    // H_0 = H(salt + password)
    auto salt_plus_password = salt;
    std::for_each(password.begin(), password.end(), [&salt_plus_password](char16_t c) {
        salt_plus_password.insert(
            salt_plus_password.end(),
            reinterpret_cast<char *>(&c),
            reinterpret_cast<char *>(&c) + sizeof(char16_t));
    });
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
        aes_ecb_decrypt(encrypted_verifier, key));
    auto decrypted_verifier_hash = aes_ecb_decrypt(
        encrypted_verifier_hash, key);
    decrypted_verifier_hash.resize(verifier_hash_size);

    if (calculated_verifier_hash != decrypted_verifier_hash)
    {
        throw xlnt::exception("bad password");
    }

    offset = 0;
    auto decrypted_size = read_int<std::uint64_t>(encrypted_package, offset);
    auto decrypted = aes_ecb_decrypt(encrypted_package, key, offset);
    decrypted.resize(static_cast<std::size_t>(decrypted_size));

    return decrypted;
}

std::vector<std::uint8_t> decrypt_xlsx_agile(
    const std::vector<std::uint8_t> &encryption_info,
    const std::u16string &password,
    const std::vector<std::uint8_t> &encrypted_package)
{
    using xlnt::detail::decode_base64;

    static const auto &xmlns = xlnt::constants::ns("encryption");
    static const auto &xmlns_p = xlnt::constants::ns("encryption-password");
    // static const auto &xmlns_c = xlnt::constants::namespace_("encryption-certificate");

    encryption_info::agile_encryption_info result;

    xml::parser parser(encryption_info.data(), encryption_info.size(), "EncryptionInfo");

    parser.next_expect(xml::parser::event_type::start_element, xmlns, "encryption");

    parser.next_expect(xml::parser::event_type::start_element, xmlns, "keyData");
    result.key_data.salt_size = parser.attribute<std::size_t>("saltSize");
    result.key_data.block_size = parser.attribute<std::size_t>("blockSize");
    result.key_data.key_bits = parser.attribute<std::size_t>("keyBits");
    result.key_data.hash_size = parser.attribute<std::size_t>("hashSize");
    result.key_data.cipher_algorithm = parser.attribute("cipherAlgorithm");
    result.key_data.cipher_chaining = parser.attribute("cipherChaining");
    result.key_data.hash_algorithm = parser.attribute("hashAlgorithm");
    result.key_data.salt_value = decode_base64(parser.attribute("saltValue"));
    parser.next_expect(xml::parser::event_type::end_element, xmlns, "keyData");

    parser.next_expect(xml::parser::event_type::start_element, xmlns, "dataIntegrity");
    result.data_integrity.hmac_key = decode_base64(parser.attribute("encryptedHmacKey"));
    result.data_integrity.hmac_value = decode_base64(parser.attribute("encryptedHmacValue"));
    parser.next_expect(xml::parser::event_type::end_element, xmlns, "dataIntegrity");

    parser.next_expect(xml::parser::event_type::start_element, xmlns, "keyEncryptors");
    parser.next_expect(xml::parser::event_type::start_element, xmlns, "keyEncryptor");
    parser.attribute("uri");
    bool any_password_key = false;

    while (parser.peek() != xml::parser::event_type::end_element)
    {
        parser.next_expect(xml::parser::event_type::start_element);

        if (parser.namespace_() == xmlns_p && parser.name() == "encryptedKey")
        {
            any_password_key = true;
            result.key_encryptor.spin_count = parser.attribute<std::size_t>("spinCount");
            result.key_encryptor.salt_size = parser.attribute<std::size_t>("saltSize");
            result.key_encryptor.block_size = parser.attribute<std::size_t>("blockSize");
            result.key_encryptor.key_bits = parser.attribute<std::size_t>("keyBits");
            result.key_encryptor.hash_size = parser.attribute<std::size_t>("hashSize");
            result.key_encryptor.cipher_algorithm = parser.attribute("cipherAlgorithm");
            result.key_encryptor.cipher_chaining = parser.attribute("cipherChaining");
            result.key_encryptor.hash = parser.attribute<xlnt::detail::hash_algorithm>("hashAlgorithm");
            result.key_encryptor.salt_value =
                decode_base64(parser.attribute("saltValue"));
            result.key_encryptor.verifier_hash_input =
                decode_base64(parser.attribute("encryptedVerifierHashInput"));
            result.key_encryptor.verifier_hash_value =
                decode_base64(parser.attribute("encryptedVerifierHashValue"));
            result.key_encryptor.encrypted_key_value =
                decode_base64(parser.attribute("encryptedKeyValue"));
        }
        else
        {
            throw xlnt::unsupported("other encryption key types not supported");
        }

        parser.next_expect(xml::parser::event_type::end_element);
    }

    if (!any_password_key)
    {
        throw xlnt::exception("no password key in keyEncryptors");
    }

    parser.next_expect(xml::parser::event_type::end_element, xmlns, "keyEncryptor");
    parser.next_expect(xml::parser::event_type::end_element, xmlns, "keyEncryptors");

    parser.next_expect(xml::parser::event_type::end_element, xmlns, "encryption");

    // begin key generation algorithm

    // H_0 = H(salt + password)
    auto salt_plus_password = result.key_encryptor.salt_value;

    std::for_each(password.begin(), password.end(), [&salt_plus_password](std::uint16_t c) {
        salt_plus_password.insert(salt_plus_password.end(), reinterpret_cast<char *>(&c),
            reinterpret_cast<char *>(&c) + sizeof(std::uint16_t));
    });

    auto h_0 = hash(result.key_encryptor.hash, salt_plus_password);

    // H_n = H(iterator + H_n-1)
    std::vector<std::uint8_t> iterator_plus_h_n(4, 0);
    iterator_plus_h_n.insert(iterator_plus_h_n.end(), h_0.begin(), h_0.end());
    std::uint32_t &iterator = *reinterpret_cast<std::uint32_t *>(iterator_plus_h_n.data());
    std::vector<std::uint8_t> h_n;
    for (iterator = 0; iterator < result.key_encryptor.spin_count; ++iterator)
    {
        hash(result.key_encryptor.hash, iterator_plus_h_n, h_n);
        std::copy(h_n.begin(), h_n.end(), iterator_plus_h_n.begin() + 4);
    }

    static const std::size_t block_size = 8;

    auto calculate_block = [&result](
        const std::vector<std::uint8_t> &raw_key,
        const std::array<std::uint8_t, block_size> &block,
        const std::vector<std::uint8_t> &encrypted)
    {
        auto combined = raw_key;
        combined.insert(combined.end(), block.begin(), block.end());

        auto key = hash(result.key_encryptor.hash, combined);
        key.resize(result.key_encryptor.key_bits / 8);

        using xlnt::detail::aes_cbc_decrypt;

        return aes_cbc_decrypt(encrypted, key, result.key_encryptor.salt_value);
    };

    const std::array<std::uint8_t, block_size> input_block_key = {{0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79}};
    auto hash_input = calculate_block(h_n, input_block_key, result.key_encryptor.verifier_hash_input);
    auto calculated_verifier = hash(result.key_encryptor.hash, hash_input);

    const std::array<std::uint8_t, block_size> verifier_block_key = {
        {0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e}};
    auto expected_verifier = calculate_block(h_n, verifier_block_key, result.key_encryptor.verifier_hash_value);
    expected_verifier.resize(calculated_verifier.size());

    if (calculated_verifier != expected_verifier)
    {
        throw xlnt::exception("bad password");
    }

    const std::array<std::uint8_t, block_size> key_value_block_key = 
    {
        {0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6}
    };
    auto key = calculate_block(h_n, key_value_block_key, result.key_encryptor.encrypted_key_value);

    auto salt_size = result.key_data.salt_size;
    auto salt_with_block_key = result.key_data.salt_value;
    salt_with_block_key.resize(salt_size + sizeof(std::uint32_t), 0);

    auto &segment = *reinterpret_cast<std::uint32_t *>(salt_with_block_key.data() + salt_size);
    auto total_size = static_cast<std::size_t>(*reinterpret_cast<const std::uint64_t *>(encrypted_package.data()));

    std::vector<std::uint8_t> encrypted_segment(POLE::OleSegmentLength, 0);
    std::vector<std::uint8_t> decrypted_package;
    decrypted_package.reserve(encrypted_package.size() - 8);

    for (std::size_t i = 8; i < encrypted_package.size(); i += POLE::OleSegmentLength)
    {
        auto iv = hash(result.key_encryptor.hash, salt_with_block_key);
        iv.resize(16);

        auto segment_begin = encrypted_package.begin() + static_cast<std::ptrdiff_t>(i);
        auto current_segment_length = std::min(POLE::OleSegmentLength, encrypted_package.size() - i);
        auto segment_end = encrypted_package.begin() + static_cast<std::ptrdiff_t>(i + current_segment_length);
        encrypted_segment.assign(segment_begin, segment_end);
        auto decrypted_segment = xlnt::detail::aes_cbc_decrypt(encrypted_segment, key, iv);
        decrypted_segment.resize(current_segment_length);

        decrypted_package.insert(
            decrypted_package.end(),
            decrypted_segment.begin(),
            decrypted_segment.end());

        ++segment;
    }

    decrypted_package.resize(total_size);

    return decrypted_package;
}

std::vector<std::uint8_t> decrypt_xlsx(
    const std::vector<std::uint8_t> &bytes,
    const std::u16string &password)
{
    if (bytes.empty())
    {
        throw xlnt::exception("empty file");
    }

    POLE::Storage storage(const_cast<std::uint8_t *>(bytes.data()), bytes.size());

    if (!storage.open())
    {
        throw xlnt::exception("not an ole compound file");
    }

    auto encrypted_package = storage.file("EncryptedPackage");
    auto encryption_info = storage.file("EncryptionInfo");

    using xlnt::detail::read_int;
    std::size_t offset = 0;

    auto version_major = read_int<std::uint16_t>(encryption_info, offset);
    auto version_minor = read_int<std::uint16_t>(encryption_info, offset);
    auto encryption_flags = read_int<std::uint32_t>(encryption_info, offset);

    // get rid of header
    encryption_info.erase(
        encryption_info.begin(),
        encryption_info.begin() + static_cast<std::ptrdiff_t>(offset));

    // version 4.4 is agile
    if (version_major == 4 && version_minor == 4)
    {
        if (encryption_flags != 0x40)
        {
            throw xlnt::exception("bad header");
        }

        return decrypt_xlsx_agile(encryption_info, password, encrypted_package);
    }

    // not agile, only try to decrypt versions 3.2 and 4.2
    if (version_minor != 2 || (version_major != 2 && version_major != 3 && version_major != 4))
    {
        throw xlnt::exception("unsupported encryption version");
    }

    if ((encryption_flags & 0b00000011) != 0) // Reserved1 and Reserved2, MUST be 0
    {
        throw xlnt::exception("bad header");
    }

    if ((encryption_flags & 0b00000100) == 0 // fCryptoAPI
        || (encryption_flags & 0b00010000) != 0) // fExternal
    {
        throw xlnt::exception("extensible encryption is not supported");
    }

    if ((encryption_flags & 0b00100000) == 0) // fAES
    {
        throw xlnt::exception("not an OOXML document");
    }

    return decrypt_xlsx_standard(encryption_info, password, encrypted_package);
}

} // namespace

namespace xlnt {
namespace detail {

std::vector<std::uint8_t> XLNT_API decrypt_xlsx(const std::vector<std::uint8_t> &data, const std::string &password)
{
    return ::decrypt_xlsx(data, utf8_to_utf16(password));
}

void xlsx_consumer::read(std::istream &source, const std::string &password)
{
    std::vector<std::uint8_t> data((std::istreambuf_iterator<char>(source)), (std::istreambuf_iterator<char>()));
    const auto decrypted = decrypt_xlsx(data, password);
    vector_istreambuf decrypted_buffer(decrypted);
    std::istream decrypted_stream(&decrypted_buffer);
    read(decrypted_stream);
}

} // namespace detail
} // namespace xlnt
