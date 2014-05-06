#include "cell.h"

namespace xlnt {

class cell_impl
{
public:
    variant value;
};

cell::cell() : impl_(nullptr)
{

}

variant cell::value()
{
    if(impl_ == nullptr)
    {
        return variant();
    }

    return impl_->value;
}

} // namespace xlnt
