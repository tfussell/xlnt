#include "cell_impl.hpp"
#include "worksheet/worksheet.hpp"

namespace xlnt {
namespace detail {

cell_impl::cell_impl() : parent_(nullptr), column_(0), row_(0), style_(nullptr), merged(false), is_date_(false), has_hyperlink_(false)
{
}
    
cell_impl::cell_impl(worksheet_impl *parent, int column_index, int row_index) : parent_(parent), column_(column_index), row_(row_index), style_(nullptr), merged(false), is_date_(false), has_hyperlink_(false)
{
}
    
cell_impl::cell_impl(const cell_impl &rhs) : is_date_(false)
{
    *this = rhs;
}
    
cell_impl &cell_impl::operator=(const cell_impl &rhs)
{
    parent_ = rhs.parent_;
    value_ = rhs.value_;
    hyperlink_ = rhs.hyperlink_;
    formula_ = rhs.formula_;
    column_ = rhs.column_;
    row_ = rhs.row_;
    style_ = rhs.style_;
    merged = rhs.merged;
    is_date_ = rhs.is_date_;
    has_hyperlink_ = rhs.has_hyperlink_;
    return *this;
}

} // namespace detail
} // namespace xlnt
