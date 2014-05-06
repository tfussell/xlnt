#include "range.h"

namespace xlnt {

class range_impl
{
public:
    cell cell(const std::string &address)
    {
        return cells_[address];
    }

private:
    std::unordered_map<std::string, class cell> cells_;
};

range::range() : impl_(nullptr)
{

}

cell range::cell(const std::string &address)
{
    if(impl_ == nullptr)
    {
        return class cell();
    }

    return impl_->cell(address);
}

} // namespace xlnt
