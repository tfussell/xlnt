#include "cell_impl.hpp"
#include "worksheet/worksheet.hpp"

namespace xlnt {
namespace detail {

  cell_impl::cell_impl() : parent_(nullptr), type_(cell::type::null), column(0), row(0), style_(nullptr), merged(false), has_hyperlink_(false)
{
}
    
cell_impl::cell_impl(worksheet_impl *parent, int column_index, int row_index) : parent_(parent), type_(cell::type::null), column(column_index), row(row_index), style_(nullptr), merged(false), has_hyperlink_(false)
{
}
    
cell_impl::cell_impl(const cell_impl &rhs)
{
    *this = rhs;
}
    
cell_impl &cell_impl::operator=(const cell_impl &rhs)
{
    parent_ = rhs.parent_;
    type_ = rhs.type_;
    numeric_value = rhs.numeric_value;
    string_value = rhs.string_value;
    hyperlink_ = rhs.hyperlink_;
    column = rhs.column;
    row = rhs.row;
    style_ = rhs.style_;
    merged = rhs.merged;
    is_date_ = rhs.is_date_;
    has_hyperlink_ = rhs.has_hyperlink_;
    return *this;
}

} // namespace detail
} // namespace xlnt
