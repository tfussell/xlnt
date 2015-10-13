#pragma once

#include <xlnt/cell/comment.hpp>

namespace xlnt {
namespace detail {

struct cell_impl;

struct comment_impl
{
    comment_impl();
    comment_impl(cell_impl *parent, const std::string &text, const std::string &author);
    comment_impl(const comment_impl &rhs);
    comment_impl &operator=(const comment_impl &rhs);
    
    std::string text_;
    std::string author_;
};

} // namespace detail
} // namespace xlnt
