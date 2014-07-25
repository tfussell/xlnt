#pragma once

#include <string>

namespace xlnt {
namespace detail {
struct cell_impl;
struct comment_impl;
} // namespace detail

class comment
{
public:
    comment(const std::string &text, const std::string &author);
    ~comment();
    std::string get_text() const;
    std::string get_author() const;

private:
    friend class cell;
    comment(detail::comment_impl *d);
    detail::comment_impl *d_;
};

} // namespace xlnt
