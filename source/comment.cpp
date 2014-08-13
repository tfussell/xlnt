#include <xlnt/cell/comment.hpp>

namespace xlnt {

comment::comment(const std::string &text, const std::string &author) : text_(text), author_(author)
{
}

comment::comment()
{
}

comment::~comment()
{
}

std::string comment::get_author() const
{
    return author_;
}

std::string comment::get_text() const
{
    return text_;
}

} // namespace xlnt
