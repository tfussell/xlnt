#include "cell.h"

namespace xlnt {

class cell_impl
{
public:
    variant value;
};

cell::cell() : impl_(std::make_shared<cell_impl>())
{

}

variant &cell::value()
{
    return impl_->value;
}

} // namespace xlnt
