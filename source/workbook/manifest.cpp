#include <xlnt/workbook/manifest.hpp>

namespace xlnt {

bool manifest::has_default_type(const std::string &extension) const
{
    return std::find_if(default_types_.begin(), default_types_.end(),
        [&](const default_type &d) { return d.get_extension() == extension; }) != default_types_.end();
}

bool manifest::has_override_type(const std::string &part_name) const
{
    return std::find_if(override_types_.begin(), override_types_.end(),
        [&](const override_type &d) { return d.get_part_name() == part_name; }) != override_types_.end();
}

std::string manifest::get_default_type(const std::string &extension) const
{
    auto match = std::find_if(default_types_.begin(), default_types_.end(),
        [&](const default_type &d) { return d.get_extension() == extension; });
    
    if(match == default_types_.end())
    {
        throw std::runtime_error("no default type found for extension: " + extension);
    }
    
    return match->get_content_type();
}

std::string manifest::get_override_type(const std::string &part_name) const
{
    auto match = std::find_if(override_types_.begin(), override_types_.end(),
        [&](const override_type &d) { return d.get_part_name() == part_name; });
    
    if(match == override_types_.end())
    {
        throw std::runtime_error("no default type found for part name: " + part_name);
    }
    
    return match->get_content_type();
}

} // namespace xlnt
