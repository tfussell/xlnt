#pragma once

#include <string>

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
    friend class worksheet;
    
    static named_range allocate(const std::string &name, worksheet ws, const range_reference &target);
    
    named_range(named_range_struct *);
    named_range_struct *root_;
};

} // namespace xlnt
