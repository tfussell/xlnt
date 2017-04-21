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

#include <detail/constants.hpp>
#include <detail/unicode.hpp>
#include <detail/cryptography/aes.hpp>
#include <detail/cryptography/base64.hpp>
#include <detail/cryptography/compound_document.hpp>
#include <detail/cryptography/encryption_info.hpp>
#include <detail/cryptography/value_traits.hpp>
#include <detail/external/include_libstudxml.hpp>
#include <detail/serialization/vector_streambuf.hpp>
#include <detail/serialization/xlsx_producer.hpp>
#include <xlnt/utils/exceptions.hpp>

namespace {

using xlnt::detail::encryption_info;

encryption_info generate_encryption_info(const std::u16string &password)
{
    encryption_info result;

    result.is_agile = true;

    result.agile = encryption_info::agile_encryption_info{};

    result.agile.key_data.block_size = 16;
    result.agile.key_data.cipher_algorithm = "AES";
    result.agile.key_data.cipher_chaining = "CBC";
    result.agile.key_data.hash_algorithm = "SHA512";
    result.agile.key_data.hash_size = 64;
    result.agile.key_data.key_bits = 128;
    result.agile.key_data.salt_size = 10;
    result.agile.key_data.salt_value = { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 };

    result.agile.data_integrity.hmac_key = { 1, 2, 3 };
    result.agile.data_integrity.hmac_value = { 3, 4, 5 };

    result.agile.key_encryptor.spin_count = 100000;
    result.agile.key_encryptor.block_size = 16;
    result.agile.key_encryptor.cipher_algorithm = "AES";
    result.agile.key_encryptor.cipher_chaining = "CBC";
    result.agile.key_encryptor.hash = xlnt::detail::hash_algorithm::sha512;
    result.agile.key_encryptor.hash_size = 64;
    result.agile.key_encryptor.key_bits = 128;
    result.agile.key_encryptor.salt_size = 10;
    result.agile.key_encryptor.salt_value = { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 };
    result.agile.key_encryptor.verifier_hash_input = { 1, 2, 3 };
    result.agile.key_encryptor.verifier_hash_value = { 1, 2, 3 };
    result.agile.key_encryptor.encrypted_key_value = { 3, 4, 5 };

    result.agile.key_data.salt_value.assign(
        reinterpret_cast<const std::uint8_t *>(password.data()),
        reinterpret_cast<const std::uint8_t *>(password.data() + password.size()));

    return result;
}

std::vector<std::uint8_t> write_agile_encryption_info(
    const encryption_info &info)
{
    static const auto &xmlns = xlnt::constants::ns("encryption");
    static const auto &xmlns_p = xlnt::constants::ns("encryption-password");

    std::vector<std::uint8_t> encryption_info;
    xlnt::detail::vector_ostreambuf encryption_info_buffer(encryption_info);
    std::ostream encryption_info_stream(&encryption_info_buffer);
    xml::serializer serializer(encryption_info_stream, "EncryptionInfo");

    serializer.start_element(xmlns, "encryption");

    const auto key_data = info.agile.key_data;
    serializer.start_element(xmlns, "keyData");
    serializer.attribute("saltSize", key_data.salt_size);
    serializer.attribute("blockSize", key_data.block_size);
    serializer.attribute("keyBits", key_data.key_bits);
    serializer.attribute("hashSize", key_data.hash_size);
    serializer.attribute("cipherAlgorithm", key_data.cipher_algorithm);
    serializer.attribute("cipherChaining", key_data.cipher_chaining);
    serializer.attribute("hashAlgorithm", key_data.hash_algorithm);
    serializer.attribute("saltValue",
        xlnt::detail::encode_base64(key_data.salt_value));
    serializer.end_element(xmlns, "keyData");

    const auto data_integrity = info.agile.data_integrity;
    serializer.start_element(xmlns, "dataIntegrity");
    serializer.attribute("encryptedHmacKey",
        xlnt::detail::encode_base64(data_integrity.hmac_key));
    serializer.attribute("encryptedHmacValue",
        xlnt::detail::encode_base64(data_integrity.hmac_value));
    serializer.end_element(xmlns, "dataIntegrity");

    const auto key_encryptor = info.agile.key_encryptor;
    serializer.start_element(xmlns, "keyEncryptors");
    serializer.start_element(xmlns, "keyEncryptor");
    serializer.attribute("uri", "");
    serializer.start_element(xmlns_p, "encryptedKey");
    serializer.attribute("spinCount", key_encryptor.spin_count);
    serializer.attribute("saltSize", key_encryptor.salt_size);
    serializer.attribute("blockSize", key_encryptor.block_size);
    serializer.attribute("keyBits", key_encryptor.key_bits);
    serializer.attribute("hashSize", key_encryptor.hash_size);
    serializer.attribute("cipherAlgorithm", key_encryptor.cipher_algorithm);
    serializer.attribute("cipherChaining", key_encryptor.cipher_chaining);
    serializer.attribute("hashAlgorithm", key_encryptor.hash);
    serializer.attribute("saltValue",
        xlnt::detail::encode_base64(key_encryptor.salt_value));
    serializer.attribute("encryptedVerifierHashInput",
        xlnt::detail::encode_base64(key_encryptor.verifier_hash_input));
    serializer.attribute("encryptedVerifierHashValue",
        xlnt::detail::encode_base64(key_encryptor.verifier_hash_value));
    serializer.attribute("encryptedKeyValue",
        xlnt::detail::encode_base64(key_encryptor.encrypted_key_value));
    serializer.end_element(xmlns_p, "encryptedKey");
    serializer.end_element(xmlns, "keyEncryptor");
    serializer.end_element(xmlns, "keyEncryptors");

    serializer.end_element(xmlns, "encryption");

    return encryption_info;
}

std::vector<std::uint8_t> write_standard_encryption_info(
    const encryption_info &/*info*/)
{
    return {};
}

std::vector<std::uint8_t> encrypt_xlsx_agile(
    const encryption_info &/*info*/,
    const std::vector<std::uint8_t> &/*plaintext*/)
{
    return {};
}

std::vector<std::uint8_t> encrypt_xlsx_standard(
    const encryption_info &/*info*/,
    const std::vector<std::uint8_t> &/*plaintext*/)
{
    //auto key = info.calculate_key();

    return {};
}

std::vector<std::uint8_t> encrypt_xlsx(
    const std::vector<std::uint8_t> &plaintext,
    const std::u16string &password)
{
    auto encryption_info = generate_encryption_info(password);

    xlnt::detail::compound_document document;

    document.add_stream("EncryptionInfo", encryption_info.is_agile
        ? write_agile_encryption_info(encryption_info)
        : write_standard_encryption_info(encryption_info));
    document.add_stream("EncryptedPackage", encryption_info.is_agile
        ? encrypt_xlsx_agile(encryption_info, plaintext)
        : encrypt_xlsx_standard(encryption_info, plaintext));

    return document.save();
}

} // namespace

namespace xlnt {
namespace detail {

std::vector<std::uint8_t> XLNT_API encrypt_xlsx(
    const std::vector<std::uint8_t> &plaintext,
    const std::string &password)
{
    return ::encrypt_xlsx(plaintext, utf8_to_utf16(password));
}

void xlsx_producer::write(std::ostream &destination, const std::string &password)
{
    std::vector<std::uint8_t> plaintext;
    vector_ostreambuf plaintext_buffer(plaintext);
    std::ostream decrypted_stream(&plaintext_buffer);
    write(decrypted_stream);

    const auto ciphertext = ::encrypt_xlsx(plaintext, utf8_to_utf16(password));
    vector_istreambuf encrypted_buffer(ciphertext);

    destination << &encrypted_buffer;
}

} // namespace detail
} // namespace xlnt
