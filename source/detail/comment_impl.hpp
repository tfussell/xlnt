#pragma once

#include "cell/cell_reference.hpp"

namespace xlnt {
namespace detail {

struct worksheet_impl;

struct comment_impl
{
    comment_impl();
    comment_impl(worksheet_impl *parent, const cell_reference &ref, const std::string &text, const std::string &author);
    comment_impl(const comment_impl &rhs);
    comment_impl &operator=(const comment_impl &rhs);
    
    worksheet_impl *parent_worksheet_;
    cell_reference parent_cell_;
    std::string text_;
    std::string author_;
};

} // namespace detail
} // namespace xlnt
