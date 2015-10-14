#include <xlnt/worksheet/worksheet.hpp>

#include "cell_impl.hpp"
#include "comment_impl.hpp"

namespace xlnt {
namespace detail {

    
    cell::type type_;
    
    worksheet_impl *parent_;
    
    column_t column_;
    row_t row_;
    
    std::string value_string_;
    long double value_numeric_;
    
    std::string formula_;
    
    bool has_hyperlink_;
    relationship hyperlink_;
    
    bool is_merged_;
    bool is_date_;
    
    std::size_t xf_index_;
    
    std::unique_ptr<style> style_;
    std::unique_ptr<comment_impl> comment_;
    
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
      is_date_(false),
      xf_index_(0),
      style_(nullptr),
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
