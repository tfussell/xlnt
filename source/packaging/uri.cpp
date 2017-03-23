#include <xlnt/packaging/uri.hpp>

namespace xlnt {

uri::uri()
{
}

uri::uri(const std::string &uri_string)
    : path_(uri_string)
{
}

std::string uri::to_string() const
{
    return path_.string();
}

path uri::path() const
{
    return path_;
}

bool uri::operator==(const uri &other) const
{
    return to_string() == other.to_string();
}

} // namespace xlnt
