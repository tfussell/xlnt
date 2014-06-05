#include "common/relationship.hpp"

namespace xlnt {

relationship::relationship(const std::string &t, const std::string &r_id, const std::string &target_uri) : id_(r_id), source_uri_(""), target_uri_(target_uri)
{
    if(t == "hyperlink")
    {
        type_ = type::hyperlink;
        target_mode_ = target_mode::external;
    }
}

} // namespace xlnt
