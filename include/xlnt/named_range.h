#pragma once

#include <string>

#include "range_reference.h"
#include "worksheet.h"

namespace xlnt {
    
class range_reference;
class worksheet;
    
struct named_range_struct;

class named_range
{
public:
    named_range();
    named_range(const std::string &name, worksheet ws, const range_reference &target);
    range_reference get_target_range() const;
    worksheet get_scope() const;
    void set_scope(worksheet scope);
    bool operator==(const named_range &comparand) const;
    
private:
    std::string name_;
    worksheet parent_worksheet_;
    range_reference target_range_;
};

} // namespace xlnt
