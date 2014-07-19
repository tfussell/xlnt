#include "common/relationship.hpp"

namespace xlnt {

relationship::relationship(type t, const std::string &r_id, const std::string &target_uri) : type_(t), id_(r_id), source_uri_(""), target_uri_(target_uri)
{
    if(t == type::hyperlink)
    {
        target_mode_ = target_mode::external;
    }
}

relationship::relationship() : type_(type::invalid),  id_(""), source_uri_(""), target_uri_("")
{
}

} // namespace xlnt
