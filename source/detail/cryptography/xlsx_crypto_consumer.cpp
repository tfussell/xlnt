// Copyright (c) 2014-2018 Thomas Fussell
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

#include <detail/binary.hpp>
#include <detail/constants.hpp>
#include <detail/unicode.hpp>
#include <detail/cryptography/encryption_info.hpp>
#include <detail/cryptography/aes.hpp>
#include <detail/cryptography/base64.hpp>
#include <detail/cryptography/compound_document.hpp>
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

std::vector<std::uint8_t> decrypt_xlsx_standard(
    encryption_info info,
    std::istream &encrypted_package_stream)
{
    const auto key = info.calculate_key();

    auto decrypted_size = read<std::uint64_t>(encrypted_package_stream);

    std::vector<std::uint8_t> encrypted_segment(4096, 0);
    std::vector<std::uint8_t> decrypted_package;

    while (encrypted_package_stream)
    {
        encrypted_package_stream.read(
            reinterpret_cast<char *>(encrypted_segment.data()),
            static_cast<std::streamsize>(encrypted_segment.size()));
        auto decrypted_segment = xlnt::detail::aes_ecb_decrypt(encrypted_segment, key);

        decrypted_package.insert(
            decrypted_package.end(),
            decrypted_segment.begin(),
            decrypted_segment.end());
    }

    decrypted_package.resize(static_cast<std::size_t>(decrypted_size));

    return decrypted_package;
}

std::vector<std::uint8_t> decrypt_xlsx_agile(
    const encryption_info &info,
    std::istream &encrypted_package_stream)
{
    const auto key = info.calculate_key();

    auto salt_size = info.agile.key_data.salt_size;
    auto salt_with_block_key = info.agile.key_data.salt_value;
    salt_with_block_key.resize(salt_size + sizeof(std::uint32_t), 0);

    auto &segment = *reinterpret_cast<std::uint32_t *>(salt_with_block_key.data() + salt_size);
    auto total_size = read<std::uint64_t>(encrypted_package_stream);

    std::vector<std::uint8_t> encrypted_segment(4096, 0);
    std::vector<std::uint8_t> decrypted_package;

    while (encrypted_package_stream)
    {
        auto iv = hash(info.agile.key_encryptor.hash, salt_with_block_key);
        iv.resize(16);

        encrypted_package_stream.read(
            reinterpret_cast<char *>(encrypted_segment.data()),
            static_cast<std::streamsize>(encrypted_segment.size()));
        auto decrypted_segment = xlnt::detail::aes_cbc_decrypt(encrypted_segment, key, iv);

        decrypted_package.insert(
            decrypted_package.end(),
            decrypted_segment.begin(),
            decrypted_segment.end());

        ++segment;
    }

    decrypted_package.resize(static_cast<std::size_t>(total_size));

    return decrypted_package;
}

encryption_info::standard_encryption_info read_standard_encryption_info(std::istream &info_stream)
{
    encryption_info::standard_encryption_info result;

    auto header_length = read<std::uint32_t>(info_stream);
    auto index_at_start = info_stream.tellg();
    /*auto skip_flags = */ read<std::uint32_t>(info_stream);
    /*auto size_extra = */ read<std::uint32_t>(info_stream);
    auto alg_id = read<std::uint32_t>(info_stream);

    if (alg_id == 0 || alg_id == 0x0000660E || alg_id == 0x0000660F || alg_id == 0x00006610)
    {
        result.cipher = xlnt::detail::cipher_algorithm::aes;
    }
    else
    {
        throw xlnt::exception("invalid cipher algorithm");
    }

    auto alg_id_hash = read<std::uint32_t>(info_stream);
    if (alg_id_hash != 0x00008004 && alg_id_hash == 0)
    {
        throw xlnt::exception("invalid hash algorithm");
    }

    result.key_bits = read<std::uint32_t>(info_stream);
    result.key_bytes = result.key_bits / 8;

    auto provider_type = read<std::uint32_t>(info_stream);
    if (provider_type != 0 && provider_type != 0x00000018)
    {
        throw xlnt::exception("invalid provider type");
    }

    read<std::uint32_t>(info_stream); // reserved 1
    if (read<std::uint32_t>(info_stream) != 0) // reserved 2
    {
        throw xlnt::exception("invalid header");
    }

    const auto csp_name_length = static_cast<std::size_t>((header_length 
        - (info_stream.tellg() - index_at_start)) / 2);
    auto csp_name = xlnt::detail::read_string<char16_t>(info_stream, csp_name_length);
    csp_name.pop_back(); // remove extraneous trailing null
    if (csp_name != u"Microsoft Enhanced RSA and AES Cryptographic Provider (Prototype)"
        && csp_name != u"Microsoft Enhanced RSA and AES Cryptographic Provider")
    {
        throw xlnt::exception("invalid cryptographic provider");
    }

    const auto salt_size = read<std::uint32_t>(info_stream);
    result.salt = xlnt::detail::read_vector<byte>(info_stream, salt_size);

    static const auto verifier_size = std::size_t(16);
    result.encrypted_verifier = xlnt::detail::read_vector<byte>(info_stream, verifier_size);

    /*const auto verifier_hash_size = */read<std::uint32_t>(info_stream);
    const auto encrypted_verifier_hash_size = std::size_t(32);
    result.encrypted_verifier_hash = xlnt::detail::read_vector<byte>(info_stream, encrypted_verifier_hash_size);

    return result;
}

encryption_info::agile_encryption_info read_agile_encryption_info(std::istream &info_stream)
{
    using xlnt::detail::decode_base64;

    static const auto &xmlns = xlnt::constants::ns("encryption");
    static const auto &xmlns_p = xlnt::constants::ns("encryption-password");
    // static const auto &xmlns_c = xlnt::constants::namespace_("encryption-certificate");

    encryption_info::agile_encryption_info result;

    xml::parser parser(info_stream, "EncryptionInfo");

    parser.next_expect(xml::parser::event_type::start_element, xmlns, "encryption");

    auto &key_data = result.key_data;
    parser.next_expect(xml::parser::event_type::start_element, xmlns, "keyData");
    key_data.salt_size = parser.attribute<std::size_t>("saltSize");
    key_data.block_size = parser.attribute<std::size_t>("blockSize");
    key_data.key_bits = parser.attribute<std::size_t>("keyBits");
    key_data.hash_size = parser.attribute<std::size_t>("hashSize");
    key_data.cipher_algorithm = parser.attribute("cipherAlgorithm");
    key_data.cipher_chaining = parser.attribute("cipherChaining");
    key_data.hash_algorithm = parser.attribute("hashAlgorithm");
    key_data.salt_value = decode_base64(parser.attribute("saltValue"));
    parser.next_expect(xml::parser::event_type::end_element, xmlns, "keyData");

    auto &data_integrity = result.data_integrity;
    parser.next_expect(xml::parser::event_type::start_element, xmlns, "dataIntegrity");
    data_integrity.hmac_key = decode_base64(parser.attribute("encryptedHmacKey"));
    data_integrity.hmac_value = decode_base64(parser.attribute("encryptedHmacValue"));
    parser.next_expect(xml::parser::event_type::end_element, xmlns, "dataIntegrity");

    auto &key_encryptor = result.key_encryptor;
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
            key_encryptor.spin_count = parser.attribute<std::size_t>("spinCount");
            key_encryptor.salt_size = parser.attribute<std::size_t>("saltSize");
            key_encryptor.block_size = parser.attribute<std::size_t>("blockSize");
            key_encryptor.key_bits = parser.attribute<std::size_t>("keyBits");
            key_encryptor.hash_size = parser.attribute<std::size_t>("hashSize");
            key_encryptor.cipher_algorithm = parser.attribute("cipherAlgorithm");
            key_encryptor.cipher_chaining = parser.attribute("cipherChaining");
            key_encryptor.hash = parser.attribute<xlnt::detail::hash_algorithm>("hashAlgorithm");
            key_encryptor.salt_value = decode_base64(parser.attribute("saltValue"));
            key_encryptor.verifier_hash_input = decode_base64(parser.attribute("encryptedVerifierHashInput"));
            key_encryptor.verifier_hash_value = decode_base64(parser.attribute("encryptedVerifierHashValue"));
            key_encryptor.encrypted_key_value = decode_base64(parser.attribute("encryptedKeyValue"));
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

    return result;
}

encryption_info read_encryption_info(std::istream &info_stream, const std::u16string &password)
{
    encryption_info info;
    
    info.password = password;

    auto version_major = read<std::uint16_t>(info_stream);
    auto version_minor = read<std::uint16_t>(info_stream);
    auto encryption_flags = read<std::uint32_t>(info_stream);
    
    info.is_agile = version_major == 4 && version_minor == 4;

    if (info.is_agile)
    {
        if (encryption_flags != 0x40)
        {
            throw xlnt::exception("bad header");
        }

        info.agile = read_agile_encryption_info(info_stream);
    }
    else
    {
        if (version_minor != 2 || (version_major != 2 && version_major != 3 && version_major != 4))
        {
            throw xlnt::exception("unsupported encryption version");
        }

        if ((encryption_flags & 0x03) != 0) // Reserved1 and Reserved2, MUST be 0
        {
            throw xlnt::exception("bad header");
        }

        if ((encryption_flags & 0x04) == 0 // fCryptoAPI
            || (encryption_flags & 0x10) != 0) // fExternal
        {
            throw xlnt::exception("extensible encryption is not supported");
        }

        if ((encryption_flags & 0x20) == 0) // fAES
        {
            throw xlnt::exception("not an OOXML document");
        }

        info.standard = read_standard_encryption_info(info_stream);
    }
    
    return info;
}

std::vector<std::uint8_t> decrypt_xlsx(
    const std::vector<std::uint8_t> &bytes,
    const std::u16string &password)
{
    if (bytes.empty())
    {
        throw xlnt::exception("empty file");
    }

    xlnt::detail::vector_istreambuf buffer(bytes);
    std::istream stream(&buffer);
    xlnt::detail::compound_document document(stream);

    auto &encryption_info_stream = document.open_read_stream("/EncryptionInfo");
    auto encryption_info = read_encryption_info(encryption_info_stream, password);

    auto &encrypted_package_stream = document.open_read_stream("/EncryptedPackage");

    return encryption_info.is_agile
        ? decrypt_xlsx_agile(encryption_info, encrypted_package_stream)
        : decrypt_xlsx_standard(encryption_info, encrypted_package_stream);
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
