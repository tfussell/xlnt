#ifdef CRYPTO_ENABLED

#include <array>
#include <pole.h>
#include <botan_all.h>
#include <include_libstudxml.hpp>

#include <detail/xlsx_consumer.hpp>
#include <xlnt/workbook/workbook.hpp>

namespace xlnt {
namespace detail {

enum class cipher_algorithm
{
	AES,
	RC2,
	RC4,
	DES,
	DESX,
	TripleDES,
	TripleDES_112
};

enum class cipher_chaining
{
	ECB,
	CBC,
	CFB
};

enum class hash_algorithm
{
	SHA1,
	SHA256,
	SHA384,
	SHA512,
	MD5,
	MD4,
	MD2,
	RIPEMD128,
	RIPEMD160,
	WHIRLPOOL
};

enum class encryption_algorithm
{
	aes_128,
	aes_192,
	aes_256,
	sha_1,
	sha_512
};

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
		std::string hash_algorithm;
		std::vector<std::uint8_t> salt_value;
		std::vector<std::uint8_t> verifier_hash_input;
		std::vector<std::uint8_t> verifier_hash_value;
		std::vector<std::uint8_t> encrypted_key_value;
	} key_encryptor;
};

struct standard_encryption_info
{
	std::vector<std::uint8_t> salt;
};

std::vector<std::uint8_t> get_file(POLE::Storage &storage, const std::string &name)
{
	POLE::Stream stream(&storage, name.c_str());
	if (stream.fail()) return {};
	std::vector<Botan::byte> bytes;

	for (std::size_t i = 0; i < stream.size(); ++i)
	{
		bytes.push_back(stream.getch());
	}

	return bytes;
}

template<typename InIter, typename OutIter>
void hash(const std::string &algorithm, InIter begin, InIter end, OutIter out)
{
	Botan::Pipe pipe(new Botan::Hash_Filter(algorithm));

	pipe.start_msg();
	std::for_each(begin, end, [&pipe](std::uint8_t b) { pipe.write(b); });
	pipe.end_msg();

	for (auto i : pipe.read_all())
	{
		*(out++) = i;
	}
}

template<typename T>
auto read_int(std::size_t &index, const std::vector<std::uint8_t> &raw_data)
{
	auto result = *reinterpret_cast<const T *>(&raw_data[index]);
	index += sizeof(T);
	return result;
};

Botan::SymmetricKey generate_standard_encryption_key(const std::vector<std::uint8_t> &raw_data,
	std::size_t offset, const std::string &password)
{
	auto header_length = read_int<std::uint32_t>(offset, raw_data);
	auto index_at_start = offset;
	auto skip_flags = read_int<std::uint32_t>(offset, raw_data);
	auto size_extra = read_int<std::uint32_t>(offset, raw_data);
	auto alg_id = read_int<std::uint32_t>(offset, raw_data);
	auto alg_hash_id = read_int<std::uint32_t>(offset, raw_data);
	auto key_bits = read_int<std::uint32_t>(offset, raw_data);
	auto provider_type = read_int<std::uint32_t>(offset, raw_data);
	/*auto reserved1 = */read_int<std::uint32_t>(offset, raw_data);
	/*auto reserved2 = */read_int<std::uint32_t>(offset, raw_data);

	const auto csp_name_length = header_length - (offset - index_at_start);
	std::vector<std::uint16_t> csp_name_wide(
		reinterpret_cast<const std::uint16_t *>(&*(raw_data.begin() + offset)),
		reinterpret_cast<const std::uint16_t *>(&*(raw_data.begin() + offset + csp_name_length)));
	std::string csp_name(csp_name_wide.begin(), csp_name_wide.end());
	offset += csp_name_length;

	const auto salt_size = read_int<std::uint32_t>(offset, raw_data);
	std::vector<std::uint8_t> salt(raw_data.begin() + offset, raw_data.begin() + offset + salt_size);
	offset += salt_size;

	static const auto verifier_size = std::size_t(16);
	std::vector<std::uint8_t> verifier_hash_input(raw_data.begin() + offset, raw_data.begin() + offset + verifier_size);
	offset += verifier_size;

	const auto verifier_hash_size = read_int<std::uint32_t>(offset, raw_data);
	std::vector<std::uint8_t> verifier_hash_value(raw_data.begin() + offset, raw_data.begin() + offset + 32);
	offset += verifier_hash_size;

	std::string hash_algorithm = "SHA-1";

	const auto key_size = key_bits / 8;
	const auto hash_out_size = hash_algorithm == "SHA-512" ? 64 : 20;

	auto salted_password = salt;

	for (auto c : password)
	{
		salted_password.push_back(c);
		salted_password.push_back(0);
	}

	std::vector<std::uint8_t> hash_result(hash_out_size, 0);
	hash(hash_algorithm, salted_password.begin(), salted_password.end(), hash_result.begin());

	std::vector<std::uint8_t> iterator_with_hash(4 + hash_out_size, 0);
	std::copy(hash_result.begin(), hash_result.end(), iterator_with_hash.begin() + 4);

	std::uint32_t &iterator = *reinterpret_cast<std::uint32_t *>(iterator_with_hash.data());
	static const std::size_t spin_count = 50000;

	for (iterator = 0; iterator < spin_count; ++iterator)
	{
		hash(hash_algorithm, iterator_with_hash.begin(), iterator_with_hash.end(), hash_result.begin());
	}

	auto hash_with_block_key = hash_result;
	std::vector<std::uint8_t> block_key(4, 0);
	hash_with_block_key.insert(hash_with_block_key.end(), block_key.begin(), block_key.end());
	hash(hash_algorithm, hash_with_block_key.begin(), hash_with_block_key.end(), hash_result.begin());

	std::vector<std::uint8_t> key = hash_result;
	key.resize(key_size);

	for (std::size_t i = 0; i < key_size; ++i)
	{
		key[i] = static_cast<std::uint8_t>(i < hash_out_size ? 0x36 ^ key[i] : 0x36);
	}

	hash(hash_algorithm, key.begin(), key.end(), key.begin());

	if (verifier_hash_value.size() <= key_size)
	{
		std::vector<std::uint8_t> first_part(key.begin(), key.begin() + key_size);

		for (int i = 0; i < key.size(); ++i)
		{
			key[i] = static_cast<std::uint8_t>(i < 20 ? 0x5C ^ key[i] : 0x5C);
		}

		hash(hash_algorithm, key.begin(), key.end(), key.begin() + key_size);
		std::copy(first_part.begin(), first_part.end(), key.begin());
	}

	key.resize(key_size, 0);

	//todo: verify here

	return Botan::SymmetricKey(key);
}

Botan::SymmetricKey generate_agile_encryption_key(const std::vector<std::uint8_t> &raw_data,
	std::size_t offset, const std::string &password)
{
	static const auto xmlns = std::string("http://schemas.microsoft.com/office/2006/encryption");
	static const auto xmlns_p = std::string("http://schemas.microsoft.com/office/2006/keyEncryptor/password");
	static const auto xmlns_c = std::string("http://schemas.microsoft.com/office/2006/keyEncryptor/certificate");

	auto from_base64 = [](const std::string &s)
	{
		Botan::Pipe base64_pipe(new Botan::Base64_Decoder());
		base64_pipe.process_msg(s);
		auto decoded = base64_pipe.read_all();
		return std::vector<std::uint8_t>(decoded.begin(), decoded.end());
	};

	agile_encryption_info result;

	xml::parser parser(raw_data.data() + offset, raw_data.size() - offset, "EncryptionInfo");

	parser.next_expect(xml::parser::event_type::start_element, xmlns, "encryption");

	parser.next_expect(xml::parser::event_type::start_element, xmlns, "keyData");
	result.key_data.salt_size = parser.attribute<std::size_t>("saltSize");
	result.key_data.block_size = parser.attribute<std::size_t>("blockSize");
	result.key_data.key_bits = parser.attribute<std::size_t>("keyBits");
	result.key_data.hash_size = parser.attribute<std::size_t>("hashSize");
	result.key_data.cipher_algorithm = parser.attribute("cipherAlgorithm");
	result.key_data.cipher_chaining = parser.attribute("cipherChaining");
	result.key_data.hash_algorithm = parser.attribute("hashAlgorithm");
	result.key_data.salt_value = from_base64(parser.attribute("saltValue"));
	parser.next_expect(xml::parser::event_type::end_element, xmlns, "keyData");

	parser.next_expect(xml::parser::event_type::start_element, xmlns, "dataIntegrity");
	result.data_integrity.hmac_key = from_base64(parser.attribute("encryptedHmacKey"));
	result.data_integrity.hmac_value = from_base64(parser.attribute("encryptedHmacValue"));
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
			result.key_encryptor.hash_algorithm = parser.attribute("hashAlgorithm");
			result.key_encryptor.salt_value = from_base64(parser.attribute("saltValue"));
			result.key_encryptor.verifier_hash_input = from_base64(parser.attribute("encryptedVerifierHashInput"));
			result.key_encryptor.verifier_hash_value = from_base64(parser.attribute("encryptedVerifierHashValue"));
			result.key_encryptor.encrypted_key_value = from_base64(parser.attribute("encryptedKeyValue"));
		}
		else
		{
			throw "other encryption key types not supported";
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

	const auto key_size = result.key_encryptor.key_bits / 8;
	const auto hash_out_size = result.key_encryptor.hash_algorithm == "SHA512" ? 64 : 20;

	auto salted_password = result.key_encryptor.salt_value;

	for (auto c : password)
	{
		salted_password.push_back(c);
		salted_password.push_back(0);
	}

	std::vector<std::uint8_t> hash_result(hash_out_size, 0);
	hash(result.key_encryptor.hash_algorithm, salted_password.begin(), salted_password.end(), hash_result.begin());

	std::vector<std::uint8_t> iterator_with_hash(4 + hash_out_size, 0);
	std::copy(hash_result.begin(), hash_result.end(), iterator_with_hash.begin() + 4);

	std::uint32_t &iterator = *reinterpret_cast<std::uint32_t *>(iterator_with_hash.data());

	for (iterator = 0; iterator < result.key_encryptor.spin_count; ++iterator)
	{
		hash(result.key_encryptor.hash_algorithm, iterator_with_hash.begin(), iterator_with_hash.end(), hash_result.begin());
	}

	auto hash_with_block_key = hash_result;
	std::vector<std::uint8_t> block_key(4, 0);
	hash_with_block_key.insert(hash_with_block_key.end(), block_key.begin(), block_key.end());
	hash(result.key_encryptor.hash_algorithm, hash_with_block_key.begin(), hash_with_block_key.end(), hash_result.begin());

	std::vector<std::uint8_t> key = hash_result;
	key.resize(key_size);

	for (std::size_t i = 0; i < key_size; ++i)
	{
		key[i] = static_cast<std::uint8_t>(i < hash_out_size ? 0x36 ^ key[i] : 0x36);
	}

	hash(result.key_encryptor.hash_algorithm, key.begin(), key.end(), key.begin());

	if (result.key_encryptor.verifier_hash_value.size() <= key_size)
	{
		std::vector<std::uint8_t> first_part(key.begin(), key.begin() + key_size);

		for (int i = 0; i < key.size(); ++i)
		{
			key[i] = static_cast<std::uint8_t>(i < 20 ? 0x5C ^ key[i] : 0x5C);
		}

		hash(result.key_encryptor.hash_algorithm, key.begin(), key.end(), key.begin() + key_size);
		std::copy(first_part.begin(), first_part.end(), key.begin());
	}

	key.resize(key_size, 0);

	//todo: verify here

	return Botan::SymmetricKey(key);
}

Botan::SymmetricKey generate_encryption_key(const std::vector<std::uint8_t> &raw_data, const std::string &password)
{
	std::size_t index = 0;

	auto version_major = read_int<std::uint16_t>(index, raw_data);
	auto version_minor = read_int<std::uint16_t>(index, raw_data);
	auto encryption_flags = read_int<std::uint32_t>(index, raw_data);

	// version 4.4 is agile
	if (version_major == 4 && version_minor == 4)
	{
		if (encryption_flags != 0x40)
		{
			throw "bad header";
		}

		return generate_agile_encryption_key(raw_data, index, password);
	}
	else
	{
		// not agile, only try to decrypt versions 3.2 and 4.2
		if (version_minor != 2
			|| (version_major != 2 && version_major != 3 && version_major != 4))
		{
			throw "unsupported encryption version";
		}

		if (encryption_flags & 0b00000011) // Reserved1 and Reserved2, MUST be 0
		{
			throw "bad header";
		}

		if ((encryption_flags & 0b00000100) != 0 // fCryptoAPI
			|| (encryption_flags & 0b00010000) == 0) // fExternal
		{
			throw "extensible encryption is not supported";
		}

		if ((encryption_flags & 0b00100000) == 0) // fAES
		{
			throw "not an OOXML document";
		}

		return generate_standard_encryption_key(raw_data, index, password);
	}
}

std::vector<std::uint8_t> decrypt_xlsx(const std::vector<std::uint8_t> &bytes, const std::string &password)
{
	std::vector<char> as_chars(bytes.begin(), bytes.end());
	POLE::Storage storage(as_chars.data(), static_cast<unsigned long>(bytes.size()));

	if (!storage.open())
	{
		throw "error";
	}

	auto key = generate_encryption_key(get_file(storage, "EncryptionInfo"), password);
	auto encrypted_package = get_file(storage, "EncryptedPackage");
	auto size = *reinterpret_cast<std::uint64_t *>(encrypted_package.data());

	Botan::InitializationVector iv;
	auto cipher_name = std::string("AES-128/ECB/NoPadding");
	auto cipher = Botan::get_cipher(cipher_name, key, iv, Botan::DECRYPTION);

	Botan::Pipe pipe(cipher);
	pipe.process_msg(encrypted_package.data() + 8,
		16 * static_cast<std::size_t>(std::ceil(size / 16)));
	Botan::secure_vector<Botan::byte> c1 = pipe.read_all(0);

	return std::vector<std::uint8_t>(c1.begin(), c1.begin() + size);
}

void xlsx_consumer::read(const std::vector<std::uint8_t> &source, const std::string &password)
{
	destination_.clear();
	auto decrypted = decrypt_xlsx(source, password);
	std::ofstream out("a.zip", std::ostream::binary);
	for (auto b : decrypted) out << b;
	out.close();
	source_.load(decrypted);
	populate_workbook();
}

void xlsx_consumer::read(std::istream &source, const std::string &password)
{
	std::vector<std::uint8_t> data((std::istreambuf_iterator<char>(source)), 
		std::istreambuf_iterator<char>());
	return read(data, password);
}

void xlsx_consumer::read(const path &source, const std::string &password)
{
	std::ifstream file_stream(source.string(), std::iostream::binary);
	return read(file_stream, password);
}

} // namespace detail
} // namespace xlnt

#endif
