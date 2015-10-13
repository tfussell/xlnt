#pragma once

#include <xlnt/cell/cell.hpp>
#include <xlnt/cell/comment.hpp>
#include <xlnt/common/types.hpp>
#include <xlnt/common/relationship.hpp>

#include "comment_impl.hpp"

namespace xlnt {

class style;

namespace detail {

struct worksheet_impl;
    
struct cell_impl
{
    cell_impl();
    cell_impl(column_t column, row_t row);
    cell_impl(worksheet_impl *parent, column_t column, row_t row);
    cell_impl(const cell_impl &rhs);
    cell_impl &operator=(const cell_impl &rhs);

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
};
    
} // namespace detail
} // namespace xlnt
