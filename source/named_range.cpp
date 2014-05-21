#include "named_range.h"
#include "range_reference.h"
#include "worksheet.h"

namespace xlnt {
    
void named_range::set_scope(worksheet scope)
{
    parent_worksheet_ = scope;
}

bool named_range::operator==(const xlnt::named_range &comparand) const
{
    return comparand.parent_worksheet_ == parent_worksheet_
    && comparand.target_range_ == target_range_;
}

named_range::named_range()
{
    
}

named_range::named_range(const std::string &name, worksheet ws, const range_reference &target)
: name_(name),
parent_worksheet_(ws),
target_range_(target)
{
}

range_reference named_range::get_target_range() const
{
    return target_range_;
}

worksheet named_range::get_scope() const
{
    return parent_worksheet_;
}
    
} // namespace xlnt
