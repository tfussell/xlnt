#pragma once

#include <xlnt/cell/comment.hpp>

namespace xlnt {
namespace detail {

struct cell_impl;

struct comment_impl
{
    comment_impl();
    comment_impl(cell_impl *parent, const string &text, const string &author);
    comment_impl(const comment_impl &rhs);
    comment_impl &operator=(const comment_impl &rhs);

    string text_;
    string author_;
};

} // namespace detail
} // namespace xlnt
