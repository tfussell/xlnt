#include <xlnt/worksheet/worksheet.hpp>

#include "cell_impl.hpp"
#include "comment_impl.hpp"

namespace xlnt {
namespace detail {
    
cell_impl::cell_impl() : cell_impl(1, 1)
{
}
    
cell_impl::cell_impl(column_t column, row_t row) : cell_impl(nullptr, column, row)
{
}
    
cell_impl::cell_impl(worksheet_impl *parent, column_t column, row_t row)
    : type_(cell::type::null),
      parent_(parent),
      column_(column),
      row_(row),
      value_numeric_(0),
      has_hyperlink_(false),
      is_merged_(false),
      xf_index_(0),
      has_style_(false),
      style_id_(0),
      comment_(nullptr)
{
}
    
cell_impl::cell_impl(const cell_impl &rhs)
{
    *this = rhs;
}
    
cell_impl &cell_impl::operator=(const cell_impl &rhs)
{
    parent_ = rhs.parent_;
    value_numeric_ = rhs.value_numeric_;
    value_string_ = rhs.value_string_;
    hyperlink_ = rhs.hyperlink_;
    formula_ = rhs.formula_;
    column_ = rhs.column_;
    row_ = rhs.row_;
    is_merged_ = rhs.is_merged_;
    has_hyperlink_ = rhs.has_hyperlink_;
    type_ = rhs.type_;
    style_id_ = rhs.style_id_;
    has_style_ = rhs.has_style_;
    
    if(rhs.comment_ != nullptr)
    {
        comment_.reset(new comment_impl(*rhs.comment_));
    }
    
    return *this;
}

} // namespace detail
} // namespace xlnt
