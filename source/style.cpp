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

style::style(const style &other)
{
}

const cell style::get_parent() const
{
    return cell(parent_);
}

std::size_t style::hash() const
{
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
    
    return seed;
}

const number_format &style::get_number_format() const
{
    return get_parent().get_number_format();
}

const border &style::get_border() const
{
    return get_parent().get_border();
}

const alignment &style::get_alignment() const
{
    return get_parent().get_alignment();
}
    
const fill &style::get_fill() const
{
    return get_parent().get_fill();
}

const font &style::get_font() const
{
    return get_parent().get_font();
}

const protection &style::get_protection() const
{
    return get_parent().get_protection();
}

bool style::pivot_button() const
{
    return get_parent().pivot_button();
}

bool style::quote_prefix() const
{
    return get_parent().quote_prefix();
}

} // namespace xlnt
