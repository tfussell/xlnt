#include "named_range.h"
#include "reference.h"
#include "worksheet.h"

namespace xlnt {

struct named_range_struct
{
    named_range_struct(const std::string &name, worksheet parent, const range_reference &reference)
    : name_(name),
    parent_worksheet_(parent),
    target_range_(reference)
    {
    }
    
    std::string name_;
    worksheet parent_worksheet_;
    range_reference target_range_;
};
    
void named_range::set_scope(worksheet scope)
{
    root_->parent_worksheet_ = scope;
}

named_range::named_range(named_range_struct *root) : root_(root)
{
    
}

bool named_range::operator==(const xlnt::named_range &comparand) const
{
    return comparand.root_->parent_worksheet_ == root_->parent_worksheet_
    && comparand.root_->target_range_ == root_->target_range_;
}

named_range::named_range() : root_(nullptr)
{
    
}

named_range::named_range(const std::string &name, worksheet ws, const range_reference &target)
{
    root_ = ws.create_named_range(name, target).root_;
}

range_reference named_range::get_target_range() const
{
    return root_->target_range_;
}

worksheet named_range::get_scope() const
{
    return root_->parent_worksheet_;
}

named_range named_range::allocate(const std::string &name, xlnt::worksheet ws, const xlnt::range_reference &target)
{
    return new named_range_struct(name, ws, target);
}
    
} // namespace xlnt
