#include "cell_impl.hpp"
#include "worksheet/worksheet.hpp"

namespace xlnt {
namespace detail {

cell_impl::cell_impl() : parent_(nullptr), type_(cell::type::null), column(0), row(0), style_(nullptr)
{
}
    
cell_impl::cell_impl(worksheet_impl *parent, int column_index, int row_index) : parent_(parent), type_(cell::type::null), column(column_index), row(row_index), style_(nullptr)
{
}

} // namespace detail
} // namespace xlnt
