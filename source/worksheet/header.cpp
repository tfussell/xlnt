#include <xlnt/worksheet/header.hpp>

namespace xlnt {

void header::set_text(const std::string &text)
{
    default_ = false;
    text_ = text;
}

void header::set_font_name(const std::string &font_name)
{
    default_ = false;
    font_name_ = font_name;
}

void header::set_font_size(std::size_t font_size)
{
    default_ = false;
    font_size_ = font_size;
}

void header::set_font_color(const std::string &font_color)
{
    default_ = false;
    font_color_ = font_color;
}

bool header::is_default() const
{
    return default_;
}

} // namespace xlnt