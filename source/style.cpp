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

border &style::get_border()
{
    return border_;
}

bool style::pivot_button()
{
    return pivot_button_;
}

bool style::quote_prefix()
{
    return quote_prefix_;
}

alignment &style::get_alignment()
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

font &style::get_font()
{
    return font_;
}

protection &style::get_protection()
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
