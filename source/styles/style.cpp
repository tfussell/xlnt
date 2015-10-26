#include <xlnt/cell/cell.hpp>
#include <xlnt/styles/font.hpp>
#include <xlnt/styles/number_format.hpp>
#include <xlnt/styles/protection.hpp>
#include <xlnt/styles/style.hpp>

namespace {

template <class T>
void hash_combine(std::size_t& seed, const T& v)
{
    std::hash<T> hasher;
    seed ^= hasher(v) + 0x9e3779b9 + (seed<<6) + (seed>>2);
}
    
}

namespace xlnt {

const color color::black = color(color::type::indexed, 0);
const color color::white = color(color::type::indexed, 0);

style::style() :
    id_(0),
    alignment_apply_(false),
    border_apply_(false),
    border_id_(0),
    fill_apply_(false),
    fill_id_(0),
    font_apply_(false),
    font_id_(0),
    number_format_apply_(false),
    number_format_id_(0),
    protection_apply_(false),
    pivot_button_(false),
    quote_prefix_(false)
{
}

style::style(const style &other) :
    id_(other.id_),
    alignment_(other.alignment_),
    alignment_apply_(other.alignment_apply_),
    border_(other.border_),
    border_apply_(other.border_apply_),
    border_id_(other.border_id_),
    fill_(other.fill_),
    fill_apply_(other.fill_apply_),
    fill_id_(other.fill_id_),
    font_(other.font_),
    font_apply_(other.font_apply_),
    font_id_(other.font_id_),
    number_format_(other.number_format_),
    number_format_apply_(other.number_format_apply_),
    number_format_id_(other.number_format_id_),
    protection_(other.protection_),
    protection_apply_(other.protection_apply_),
    pivot_button_(other.pivot_button_),
    quote_prefix_(other.quote_prefix_)
{
}

style &style::operator=(const style &other)
{
    id_ = other.id_;
    alignment_ = other.alignment_;
    alignment_apply_ = other.alignment_apply_;
    border_ = other.border_;
    border_apply_ = other.border_apply_;
    border_id_ = other.border_id_;
    fill_ = other.fill_;
    fill_apply_ = other.fill_apply_;
    fill_id_ = other.fill_id_;
    font_ = other.font_;
    font_apply_ = other.font_apply_;
    font_id_ = other.font_id_;
    number_format_ = other.number_format_;
    number_format_apply_ = other.number_format_apply_;
    number_format_id_ = other.number_format_id_;
    protection_ = other.protection_;
    protection_apply_ = other.protection_apply_;
    pivot_button_ = other.pivot_button_;
    quote_prefix_ = other.quote_prefix_;
    
    return *this;
}

std::size_t style::hash() const
{
    std::size_t seed = 0;
    if(alignment_apply_) hash_combine(seed, alignment_apply_);
    hash_combine(seed, alignment_.hash());
    hash_combine(seed, border_apply_);
    hash_combine(seed, border_id_);
    hash_combine(seed, font_apply_);
    hash_combine(seed, font_id_);
    hash_combine(seed, fill_apply_);
    hash_combine(seed, fill_id_);
    hash_combine(seed, number_format_apply_);
    hash_combine(seed, number_format_id_);
    hash_combine(seed, protection_apply_);
    if(protection_apply_) hash_combine(seed, get_protection().hash());
//    hash_combine(seed, pivot_button_);
//    hash_combine(seed, quote_prefix_);
    
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
