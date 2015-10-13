#include <xlnt/worksheet/worksheet.hpp>

#include "cell_impl.hpp"
#include "comment_impl.hpp"

namespace xlnt {
namespace detail {

cell_impl::cell_impl() : parent_(nullptr), type_(cell::type::null),
    column_(1), row_(1), style_(nullptr), is_merged_(false),
    is_date_(false), has_hyperlink_(false), xf_index_(0), comment_(nullptr)
{
}
    
cell_impl::cell_impl(column_t column, row_t row)
    : parent_(nullptr), type_(cell::type::null), column_(column), row_(row),
    style_(nullptr), is_merged_(false), is_date_(false),
    has_hyperlink_(false), xf_index_(0), comment_(nullptr)
    {
    }
    
cell_impl::cell_impl(worksheet_impl *parent, column_t column, row_t row)
    : parent_(parent), type_(cell::type::null), column_(column), row_(row), style_(nullptr),
    is_merged_(false), is_date_(false), has_hyperlink_(false), xf_index_(0), comment_(nullptr)
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
    is_date_ = rhs.is_date_;
    has_hyperlink_ = rhs.has_hyperlink_;
    type_ = rhs.type_;
    
    if(rhs.style_ != nullptr)
    {
        style_.reset(new style(*rhs.style_));
    }
    
    if(rhs.comment_ != nullptr)
    {
        comment_.reset(new comment_impl(*rhs.comment_));
    }
    
    return *this;
}

} // namespace detail
} // namespace xlnt
