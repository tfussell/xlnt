#pragma once

#include "cell/cell.hpp"
#include "common/types.hpp"
#include "common/relationship.hpp"

namespace xlnt {

class style;

namespace detail {

struct worksheet_impl;
    
struct cell_impl
{
    cell_impl();
    cell_impl(worksheet_impl *parent, int column_index, int row_index);
    cell_impl(const cell_impl &rhs);
    cell_impl &operator=(const cell_impl &rhs);

    worksheet_impl *parent_;
    cell::type type_;
    long double numeric_value;
    std::string string_value;
    relationship hyperlink_;
    column_t column;
    row_t row;
    style *style_;
    bool merged;
    bool is_date_;
    bool has_hyperlink_;
    comment comment_;
};
    
} // namespace detail
} // namespace xlnt
