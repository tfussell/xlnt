#pragma once

#include <cstdint>
#include <vector>

namespace xlnt {

std::vector<std::uint8_t> xaes_ecb_encrypt(
    const std::vector<std::uint8_t> &input,
    const std::vector<std::uint8_t> &key);

std::vector<std::uint8_t> xaes_ecb_decrypt(
    const std::vector<std::uint8_t> &input,
    const std::vector<std::uint8_t> &key);

std::vector<std::uint8_t> xaes_cbc_encrypt(
    const std::vector<std::uint8_t> &input,
    const std::vector<std::uint8_t> &key,
    const std::vector<std::uint8_t> &iv);

std::vector<std::uint8_t> xaes_cbc_decrypt(
    const std::vector<std::uint8_t> &input,
    const std::vector<std::uint8_t> &key,
    const std::vector<std::uint8_t> &iv);

} // namespace xlnt
