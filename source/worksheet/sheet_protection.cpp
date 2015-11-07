#include <algorithm>
#include <locale>
#include <sstream>

#include <xlnt/worksheet/sheet_protection.hpp>

namespace xlnt {

void sheet_protection::set_password(const string &password)
{
    hashed_password_ = hash_password(password);
}

template <typename T>
string int_to_hex(T i)
{
    std::stringstream stream;
    stream << std::hex << i;
    return stream.str().c_str();
}

string sheet_protection::hash_password(const string &plaintext_password)
{
    int password = 0x0000;

	for (int i = 1; i <= plaintext_password.num_bytes(); i++)
    {
		char character = plaintext_password.data()[i - 1];
        int value = character << i;
        int rotated_bits = value >> 15;
        value &= 0x7fff;
        password ^= (value | rotated_bits);
        i++;
    }

    password ^= plaintext_password.num_bytes();
    password ^= 0xCE4B;

    string hashed = int_to_hex(password);
	auto iter = hashed.data();

	while (iter != hashed.data() + hashed.num_bytes())
	{
		*iter = std::toupper(*iter, std::locale::classic());
	}

    return hashed;
}
}
