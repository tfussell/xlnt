#include "cell_impl.h"

namespace xlnt {
namespace detail {

cell_impl::cell_impl() : type_(cell::type::null), column(0), row(0)
{
}
    
cell_impl::cell_impl(int column_index, int row_index) : type_(cell::type::null), column(column_index), row(row_index)
{
}

} // namespace detail
} // namespace xlnt
