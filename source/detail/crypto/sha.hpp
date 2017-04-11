#pragma once

#include <vector>

class SHA
{
public:
    static std::vector<std::uint8_t> sha1(const std::vector<std::uint8_t> &data);
    static std::vector<std::uint8_t> sha512(const std::vector<std::uint8_t> &data);    
};
