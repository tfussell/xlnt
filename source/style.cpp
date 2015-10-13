#include <xlnt/styles/style.hpp>

namespace xlnt {

style::style(const style &rhs) : number_format_(rhs.number_format_), protection_(rhs.protection_)
{
}

void style::set_protection(xlnt::protection protection)
{
    protection_ = protection;
}
    
void style::set_number_format(xlnt::number_format format)
{
    number_format_ = format;
}

border style::get_border() const
{
    return border_;
}

bool style::pivot_button() const
{
    return pivot_button_;
}

bool style::quote_prefix() const
{
    return quote_prefix_;
}

alignment style::get_alignment() const
{
    return alignment_;
}
    
fill &style::get_fill()
{
    return fill_;
}

const fill &style::get_fill() const
{
    return fill_;
}

font style::get_font() const
{
    return font_;
}

protection style::get_protection() const
{
    return protection_;
}

void style::set_pivot_button(bool pivot)
{
    pivot_button_ = pivot;
}

void style::set_quote_prefix(bool quote)
{
    quote_prefix_ = quote;
}

} // namespace xlnt
