// Copyright (c) 2014-2016 Thomas Fussell
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
#include <pole.h>
#include <botan_all.h>
#include <include_libstudxml.hpp>

#include <detail/vector_streambuf.hpp>
#include <detail/xlsx_consumer.hpp>
#include <xlnt/utils/exceptions.hpp>
#include <xlnt/workbook/workbook.hpp>

namespace xlnt {
namespace detail {

struct crypto_helper
{
    static const std::size_t segment_length = 4096;

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

    static std::vector<std::uint8_t> aes(const std::vector<std::uint8_t> &key,
        const std::vector<std::uint8_t> &iv,
        const std::vector<std::uint8_t> &encrypted,
        cipher_chaining chaining, cipher_direction direction)
    {
        std::string cipher_name("AES-");
        cipher_name.append(std::to_string(key.size() * 8));
        cipher_name.append(chaining == cipher_chaining::ecb
            ? "/ECB/NoPadding" : "/CBC/NoPadding");

        auto botan_direction = direction == cipher_direction::decryption
            ? Botan::DECRYPTION : Botan::ENCRYPTION;
        Botan::Pipe pipe(Botan::get_cipher(cipher_name, key, iv, botan_direction));
        pipe.process_msg(encrypted);
        auto decrypted = pipe.read_all();

        return std::vector<std::uint8_t>(decrypted.begin(), decrypted.end());
    }

    static std::vector<std::uint8_t> decode_base64(const std::string &encoded)
    {
        Botan::Pipe pipe(new Botan::Base64_Decoder);
        pipe.process_msg(encoded);
        auto decoded = pipe.read_all();

        return std::vector<std::uint8_t>(decoded.begin(), decoded.end());
    };

    static std::vector<std::uint8_t> hash(hash_algorithm algorithm,
        const std::vector<std::uint8_t> &input)
    {
        Botan::Pipe pipe(new Botan::Hash_Filter(
            algorithm == hash_algorithm::sha512 ? "SHA-512" : "SHA-1"));
        pipe.process_msg(input);
        auto hash = pipe.read_all();

        return std::vector<std::uint8_t>(hash.begin(), hash.end());
    };

    static std::vector<std::uint8_t> get_file(POLE::Storage &storage, const std::string &name)
    {
        POLE::Stream stream(&storage, name.c_str());
        if (stream.fail()) return {};
        std::vector<std::uint8_t> bytes(stream.size(), 0);
        stream.read(bytes.data(), static_cast<unsigned long>(bytes.size()));
        return bytes;
    }

    template<typename T>
    static auto read_int(std::size_t &index, const std::vector<std::uint8_t> &raw_data)
    {
        auto result = *reinterpret_cast<const T *>(&raw_data[index]);
        index += sizeof(T);

        return result;
    };

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
        const std::string &password, const std::vector<std::uint8_t> &encrypted_package)
    {
        std::size_t offset = 0;

        standard_encryption_info info;

        auto header_length = read_int<std::uint32_t>(offset, encryption_info);
        auto index_at_start = offset;
        /*auto skip_flags = */read_int<std::uint32_t>(offset, encryption_info);
        /*auto size_extra = */read_int<std::uint32_t>(offset, encryption_info);
        auto alg_id = read_int<std::uint32_t>(offset, encryption_info);

        if (alg_id == 0 || alg_id == 0x0000660E || alg_id == 0x0000660F || alg_id == 0x00006610)
        {
            info.cipher = cipher_algorithm::aes;
        }
        else
        {
            throw xlnt::exception("invalid cipher algorithm");
        }

        auto alg_id_hash = read_int<std::uint32_t>(offset, encryption_info);
        if (alg_id_hash != 0x00008004 && alg_id_hash == 0)
        {
            throw xlnt::exception("invalid hash algorithm");
        }
        
        info.key_bits = read_int<std::uint32_t>(offset, encryption_info);
        info.key_bytes = info.key_bits / 8;

        auto provider_type = read_int<std::uint32_t>(offset, encryption_info);
        if (provider_type != 0 && provider_type != 0x00000018)
        {
            throw xlnt::exception("invalid provider type");
        }

        read_int<std::uint32_t>(offset, encryption_info); // reserved 1
        if (read_int<std::uint32_t>(offset, encryption_info) != 0) // reserved 2
        {
            throw xlnt::exception("invalid header");
        }

        const auto csp_name_length = header_length - (offset - index_at_start);
        std::vector<std::uint16_t> csp_name_wide(
            reinterpret_cast<const std::uint16_t *>(&*(encryption_info.begin() + static_cast<std::ptrdiff_t>(offset))),
            reinterpret_cast<const std::uint16_t *>(&*(encryption_info.begin() + static_cast<std::ptrdiff_t>(offset + csp_name_length))));
        std::string csp_name(csp_name_wide.begin(), csp_name_wide.end() - 1); // without trailing null
        if (csp_name != "Microsoft Enhanced RSA and AES Cryptographic Provider (Prototype)"
            && csp_name != "Microsoft Enhanced RSA and AES Cryptographic Provider")
        {
            throw xlnt::exception("invalid cryptographic provider");
        }
        offset += csp_name_length;

        const auto salt_size = read_int<std::uint32_t>(offset, encryption_info);
        std::vector<std::uint8_t> salt(encryption_info.begin() + static_cast<std::ptrdiff_t>(offset),
            encryption_info.begin() + static_cast<std::ptrdiff_t>(offset + salt_size));
        offset += salt_size;

        static const auto verifier_size = std::size_t(16);
        std::vector<std::uint8_t> verifier_hash_input(encryption_info.begin() + static_cast<std::ptrdiff_t>(offset),
            encryption_info.begin() + static_cast<std::ptrdiff_t>(offset + verifier_size));
        offset += verifier_size;

        const auto verifier_hash_size = read_int<std::uint32_t>(offset, encryption_info);
        std::vector<std::uint8_t> verifier_hash_value(encryption_info.begin() + static_cast<std::ptrdiff_t>(offset),
            encryption_info.begin() + static_cast<std::ptrdiff_t>(offset + verifier_hash_size));
        offset += verifier_hash_size;

        // begin key generation algorithm

        // H_0 = H(salt + password)
        auto salt_plus_password = salt;
        std::vector<std::uint16_t> password_wide(password.begin(), password.end());
        std::for_each(password_wide.begin(), password_wide.end(), 
            [&salt_plus_password](std::uint16_t c)
        {
            salt_plus_password.insert(salt_plus_password.end(), 
                reinterpret_cast<char *>(&c), 
                reinterpret_cast<char *>(&c) + sizeof(std::uint16_t));
        });
        std::vector<std::uint8_t> h_0 = hash(info.hash, salt_plus_password);

        // H_n = H(iterator + H_n-1)
        std::vector<std::uint8_t> iterator_plus_h_n(4, 0);
        iterator_plus_h_n.insert(iterator_plus_h_n.end(), h_0.begin(), h_0.end());
        std::uint32_t &iterator = *reinterpret_cast<std::uint32_t *>(iterator_plus_h_n.data());
        std::vector<std::uint8_t> h_n;
        for (iterator = 0; iterator < info.spin_count; ++iterator)
        {
            h_n = hash(info.hash, iterator_plus_h_n);
            std::copy(h_n.begin(), h_n.end(), iterator_plus_h_n.begin() + 4);
        }

        // H_final = H(H_n + block)
        auto h_n_plus_block = h_n;
        const std::uint32_t block_number = 0;
        h_n_plus_block.insert(h_n_plus_block.end(),
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
        
        auto key_derived = std::vector<std::uint8_t>(X3.begin(), X3.begin() + static_cast<std::ptrdiff_t>(info.key_bytes));

        //todo: verify here

        std::size_t package_offset = 0;
        auto decrypted_size = read_int<std::uint64_t>(package_offset, encrypted_package);
        auto decrypted = aes(key_derived, {}, std::vector<std::uint8_t>(
            encrypted_package.begin() + 8, encrypted_package.end()),
            cipher_chaining::ecb, cipher_direction::decryption);
        decrypted.resize(decrypted_size);

        return decrypted;
    }

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

    static std::vector<std::uint8_t> decrypt_xlsx_agile(const std::vector<std::uint8_t> &encryption_info,
        const std::string &password, const std::vector<std::uint8_t> &encrypted_package)
    {
        static const auto xmlns = std::string("http://schemas.microsoft.com/office/2006/encryption");
        static const auto xmlns_p = std::string("http://schemas.microsoft.com/office/2006/keyEncryptor/password");
        static const auto xmlns_c = std::string("http://schemas.microsoft.com/office/2006/keyEncryptor/certificate");

        agile_encryption_info result;

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

                auto hash_algorithm_string = parser.attribute("hashAlgorithm");
                if (hash_algorithm_string == "SHA512")
                {
                    result.key_encryptor.hash = hash_algorithm::sha512;
                }
                else if (hash_algorithm_string == "SHA1")
                {
                    result.key_encryptor.hash = hash_algorithm::sha1;
                }
                else if (hash_algorithm_string == "SHA256")
                {
                    result.key_encryptor.hash = hash_algorithm::sha256;
                }
                else if (hash_algorithm_string == "SHA384")
                {
                    result.key_encryptor.hash = hash_algorithm::sha384;
                }

                result.key_encryptor.salt_value = decode_base64(parser.attribute("saltValue"));
                result.key_encryptor.verifier_hash_input = decode_base64(parser.attribute("encryptedVerifierHashInput"));
                result.key_encryptor.verifier_hash_value = decode_base64(parser.attribute("encryptedVerifierHashValue"));
                result.key_encryptor.encrypted_key_value = decode_base64(parser.attribute("encryptedKeyValue"));
            }
            else
            {
                throw xlnt::unsupported("other encryption key types not supported");
            }

            parser.next_expect(xml::parser::event_type::end_element);
        }

        if (!any_password_key)
        {
            throw "no password key in keyEncryptors";
        }

        parser.next_expect(xml::parser::event_type::end_element, xmlns, "keyEncryptor");
        parser.next_expect(xml::parser::event_type::end_element, xmlns, "keyEncryptors");

        parser.next_expect(xml::parser::event_type::end_element, xmlns, "encryption");


        // begin key generation algorithm

        // H_0 = H(salt + password)
        auto salt_plus_password = result.key_encryptor.salt_value;
        std::vector<std::uint16_t> password_wide(password.begin(), password.end());

        std::for_each(password_wide.begin(), password_wide.end(),
            [&salt_plus_password](std::uint16_t c)
        {
            salt_plus_password.insert(salt_plus_password.end(),
                reinterpret_cast<char *>(&c),
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
            h_n = hash(result.key_encryptor.hash, iterator_plus_h_n);
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
            return aes(key, result.key_encryptor.salt_value, encrypted,
                cipher_chaining::cbc, cipher_direction::decryption);
        };

        const std::array<std::uint8_t, block_size> input_block_key 
            = { {0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79} };
        auto hash_input = calculate_block(h_n, input_block_key,
            result.key_encryptor.verifier_hash_input);
        auto calculated_verifier = hash(result.key_encryptor.hash, hash_input);

        const std::array<std::uint8_t, block_size> verifier_block_key 
            = { {0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e} };
        auto expected_verifier = calculate_block(h_n, verifier_block_key,
            result.key_encryptor.verifier_hash_value);

        if (calculated_verifier.size() != expected_verifier.size()
            || std::mismatch(calculated_verifier.begin(), calculated_verifier.end(),
                expected_verifier.begin(), expected_verifier.end())
                != std::make_pair(calculated_verifier.end(), expected_verifier.end()))
        {
            throw xlnt::exception("bad password");
        }

        const std::array<std::uint8_t, block_size> key_value_block_key 
            = { {0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6} };
        auto key = calculate_block(h_n, key_value_block_key, 
            result.key_encryptor.encrypted_key_value);

        auto salt_size = result.key_data.salt_size;
        auto salt_with_block_key = result.key_data.salt_value;
        salt_with_block_key.resize(salt_size + sizeof(std::uint32_t), 0);

        auto &segment = *reinterpret_cast<std::uint32_t *>(salt_with_block_key.data() + salt_size);
        auto total_size = static_cast<std::size_t>(*reinterpret_cast<const std::uint64_t *>(encrypted_package.data()));

        std::vector<std::uint8_t> encrypted_segment(segment_length, 0);
        std::vector<std::uint8_t> decrypted_package;
        decrypted_package.reserve(encrypted_package.size() - 8);

        for (std::size_t i = 8; i < encrypted_package.size(); i += segment_length)
        {
            auto iv = hash(result.key_encryptor.hash, salt_with_block_key);
            iv.resize(16);

            auto decrypted_segment = aes(key, iv, std::vector<std::uint8_t>(
                encrypted_package.begin() + static_cast<std::ptrdiff_t>(i),
                encrypted_package.begin() + static_cast<std::ptrdiff_t>(i) + segment_length),
                cipher_chaining::cbc, cipher_direction::decryption);

            decrypted_package.insert(decrypted_package.end(), 
                decrypted_segment.begin(), decrypted_segment.end());

            ++segment;
        }

        decrypted_package.resize(total_size);

        return decrypted_package;
    }

    static std::vector<std::uint8_t> decrypt_xlsx(const std::vector<std::uint8_t> &bytes, const std::string &password)
    {
        if (bytes.empty())
        {
            throw xlnt::exception("empty file");
        }

        std::vector<char> as_chars(bytes.begin(), bytes.end());
        POLE::Storage storage(as_chars.data(), static_cast<unsigned long>(bytes.size()));

        if (!storage.open())
        {
            throw xlnt::exception("not an ole compound file");
        }

        auto encrypted_package = get_file(storage, "EncryptedPackage");
        auto encryption_info = get_file(storage, "EncryptionInfo");

        std::size_t index = 0;

        auto version_major = read_int<std::uint16_t>(index, encryption_info);
        auto version_minor = read_int<std::uint16_t>(index, encryption_info);
        auto encryption_flags = read_int<std::uint32_t>(index, encryption_info);

        // get rid of header
        encryption_info.erase(encryption_info.begin(),
            encryption_info.begin() + static_cast<std::ptrdiff_t>(index));

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
        if (version_minor != 2
            || (version_major != 2 && version_major != 3 && version_major != 4))
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
};

void xlsx_consumer::read(std::istream &source, const std::string &password)
{
	std::vector<std::uint8_t> data((std::istreambuf_iterator<char>(source)),
		(std::istreambuf_iterator<char>()));
	const auto decrypted = crypto_helper::decrypt_xlsx(data, password);
	vector_istreambuf decrypted_buffer(decrypted);
	std::istream decrypted_stream(&decrypted_buffer);
	read(decrypted_stream);
}

} // namespace detail
} // namespace xlnt
