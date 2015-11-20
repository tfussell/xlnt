#include <xlnt/packaging/override_type.hpp>

namespace xlnt {

override_type::override_type()
{
}

override_type::override_type(const std::string &part_name, const std::string &content_type)
    : part_name_(part_name), content_type_(content_type)
{
}

override_type::override_type(const override_type &other)
    : part_name_(other.part_name_), content_type_(other.content_type_)
{
}

override_type &override_type::operator=(const override_type &other)
{
    part_name_ = other.part_name_;
    content_type_ = other.content_type_;

    return *this;
}

std::string override_type::get_part_name() const
{
    return part_name_;
}

std::string override_type::get_content_type() const
{
    return content_type_;
}

} // namespace xlnt
