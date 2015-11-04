#pragma once

#include <functional>

#include "xlnt_config.hpp"

namespace xlnt {

// TODO: Can this be in source->detail, or must it be in a header?
template <class T>
inline void XLNT_FUNCTION hash_combine(std::size_t &seed, const T &v)
{
    std::hash<T> hasher;
    seed ^= hasher(v) + 0x9e3779b9 + (seed << 6) + (seed >> 2);
}

} // namespace xlnt
