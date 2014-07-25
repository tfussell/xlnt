#include "cell/comment.hpp"
#include "detail/comment_impl.hpp"

namespace xlnt {
namespace detail {

detail::comment_impl::comment_impl() : parent_worksheet_(nullptr)
{
}

comment_impl::comment_impl(worksheet_impl *ws, const cell_reference &ref, const std::string &text, const std::string &author) : parent_worksheet_(ws), parent_cell_(ref), text_(text), author_(author)
{

}

} // namespace detail

comment::comment(const std::string &text, const std::string &author) : d_(new detail::comment_impl())
{
    d_->text_ = text;
    d_->author_ = author;
}

comment::comment(detail::comment_impl *d) : d_(d)
{
}

comment::~comment()
{
    if(d_->parent_worksheet_ == nullptr)
    {
        delete d_;
        d_ = nullptr;
    }
}

std::string comment::get_author() const
{
    return d_->author_;
}

std::string comment::get_text() const
{
    return d_->text_;
}

} // namespace xlnt
