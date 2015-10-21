#include <xlnt/cell/cell.hpp>
#include <xlnt/styles/style.hpp>

#include "font.hpp"
#include "number_format.hpp"
#include "protection.hpp"

namespace {

template <class T>
void hash_combine(std::size_t& seed, const T& v)
{
    std::hash<T> hasher;
    seed ^= hasher(v) + 0x9e3779b9 + (seed<<6) + (seed>>2);
}
    
}

namespace xlnt {

style::style()
{
}

std::size_t style::hash() const
{
    /*
    std::size_t seed = 100;
    std::hash<std::string> string_hash;
    auto font_ = get_font();
    hash_combine(seed, string_hash(font_.name_));
    hash_combine(seed, font_.size_);
    hash_combine(seed, static_cast<std::size_t>(font_.underline_));
    std::size_t bools = font_.bold_;
    bools = bools << 1 & font_.italic_;
    bools = bools << 1 & font_.strikethrough_;
    bools = bools << 1 & font_.superscript_;
    bools = bools << 1 & font_.subscript_;
    hash_combine(seed, bools);
     */
    std::size_t seed = 100;
    hash_combine(seed, number_format_.hash());
    
    return seed;
}

const number_format style::get_number_format() const
{
    return number_format_;
}

const border style::get_border() const
{
    return border_;
}

const alignment style::get_alignment() const
{
    return alignment_;
}
    
const fill style::get_fill() const
{
    return fill_;
}

const font style::get_font() const
{
    return font_;
}

const protection style::get_protection() const
{
    return protection_;
}

bool style::pivot_button() const
{
    return pivot_button_;
}

bool style::quote_prefix() const
{
    return quote_prefix_;
}

} // namespace xlnt
