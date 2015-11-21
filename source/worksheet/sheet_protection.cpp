#include <algorithm>
#include <locale>
#include <sstream>

#include <xlnt/worksheet/sheet_protection.hpp>

namespace {

template <typename T>
std::string int_to_hex(T i)
{
    std::stringstream stream;
    stream << std::hex << i;
    return stream.str();
}

} // namespace

namespace xlnt {

void sheet_protection::set_password(const std::string &password)
{
    hashed_password_ = hash_password(password);
}

std::string sheet_protection::get_hashed_password() const
{
    return hashed_password_;
}

std::string sheet_protection::hash_password(const std::string &plaintext_password)
{
    int password = 0x0000;
    int i = 1;

    for (auto character : plaintext_password)
    {
        int value = character << i;
        int rotated_bits = value >> 15;
        value &= 0x7fff;
        password ^= (value | rotated_bits);
        i++;
    }

    password ^= plaintext_password.size();
    password ^= 0xCE4B;

    std::string hashed = int_to_hex(password);
    std::transform(hashed.begin(), hashed.end(), hashed.begin(),
                   [](char c) { return std::toupper(c, std::locale::classic()); });
    return hashed;
}
}
