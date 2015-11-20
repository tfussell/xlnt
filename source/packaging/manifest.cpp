#include <algorithm>
#include <stdexcept>

#include <xlnt/packaging/manifest.hpp>

namespace {

bool match_path(const std::string &path, const std::string &comparand)
{
    if (path == comparand)
    {
        return true;
    }

    if (path[0] != '/' && path[0] != '.' && comparand[0] == '/')
    {
        return match_path("/" + path, comparand);
    }
    else if (comparand[0] != '/' && comparand[0] != '.' && path[0] == '/')
    {
        return match_path(path, "/" + comparand);
    }

    return false;
}

} // namespace

namespace xlnt {

bool manifest::has_default_type(const std::string &extension) const
{
    return std::find_if(default_types_.begin(), default_types_.end(),
                        [&](const default_type &d) { return d.get_extension() == extension; }) != default_types_.end();
}

bool manifest::has_override_type(const std::string &part_name) const
{
    return std::find_if(override_types_.begin(), override_types_.end(), [&](const override_type &d) {
               return match_path(d.get_part_name(), part_name);
           }) != override_types_.end();
}

void manifest::add_default_type(const std::string &extension, const std::string &content_type)
{
    if(has_default_type(extension)) return;
    default_types_.push_back(default_type(extension, content_type));
}

void manifest::add_override_type(const std::string &part_name, const std::string &content_type)
{
    if(has_override_type(part_name)) return;
    override_types_.push_back(override_type(part_name, content_type));
}

std::string manifest::get_default_type(const std::string &extension) const
{
    auto match = std::find_if(default_types_.begin(), default_types_.end(),
                              [&](const default_type &d) { return d.get_extension() == extension; });

    if (match == default_types_.end())
    {
        throw std::runtime_error("no default type found for extension: " + extension);
    }

    return match->get_content_type();
}

std::string manifest::get_override_type(const std::string &part_name) const
{
    auto match = std::find_if(override_types_.begin(), override_types_.end(),
                              [&](const override_type &d) { return match_path(d.get_part_name(), part_name); });

    if (match == override_types_.end())
    {
        throw std::runtime_error("no default type found for part name: " + part_name);
    }

    return match->get_content_type();
}

const std::vector<default_type> &manifest::get_default_types() const
{
    return default_types_;
}

const std::vector<override_type> &manifest::get_override_types() const
{
    return override_types_;
}

} // namespace xlnt
