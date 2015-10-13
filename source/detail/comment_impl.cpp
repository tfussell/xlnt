#include "comment_impl.hpp"


namespace xlnt {
namespace detail {

comment_impl::comment_impl()
{
}
    
comment_impl::comment_impl(const comment_impl &rhs)
{
    *this = rhs;
}

comment_impl &comment_impl::operator=(const xlnt::detail::comment_impl &rhs)
{
    text_ = rhs.text_;
    author_ = rhs.author_;
    
    return *this;
}
    
} // namespace detail
} // namespace xlnt