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
    
} // namespace xlnt
