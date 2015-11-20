#include <xlnt/packaging/default_type.hpp>

namespace xlnt {

default_type::default_type()
{
}

default_type::default_type(const std::string &extension, const std::string &content_type)
    : extension_(extension), content_type_(content_type)
{
}

default_type::default_type(const default_type &other) : extension_(other.extension_), content_type_(other.content_type_)
{
}

default_type &default_type::operator=(const default_type &other)
{
    extension_ = other.extension_;
    content_type_ = other.content_type_;

    return *this;
}

std::string default_type::get_extension() const
{
    return extension_;
}
std::string default_type::get_content_type() const
{
    return content_type_;
}

} // namespace xlnt
