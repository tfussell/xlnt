#include <xlnt/cell/comment.hpp>
#include <xlnt/cell/cell.hpp>

#include "detail/comment_impl.hpp"

namespace xlnt {

comment::comment(detail::comment_impl *d) : d_(d)
{
}
    
comment::comment(cell parent, const std::string &text, const std::string &author) : d_(nullptr)
{
    d_ = parent.get_comment().d_;
}

comment::comment() : d_(nullptr)
{
}

comment::~comment()
{
}

std::string comment::get_author() const
{
    return d_->author_;
}

std::string comment::get_text() const
{
    return d_->text_;
}
    
    bool comment::operator==(const xlnt::comment &other) const
    {
        return d_ == other.d_;
    }

} // namespace xlnt
