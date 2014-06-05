#include "cell_impl.hpp"

namespace xlnt {
namespace detail {

cell_impl::cell_impl() : type_(cell::type::null), column(0), row(0), style_(nullptr)
{
}
    
cell_impl::cell_impl(int column_index, int row_index) : type_(cell::type::null), column(column_index), row(row_index), style_(nullptr)
{
}

} // namespace detail
} // namespace xlnt
