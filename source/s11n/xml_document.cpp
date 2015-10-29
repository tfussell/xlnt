#include <xlnt/s11n/xml_document.hpp>

namespace xlnt {

void xml_document::set_encoding(const std::string &encoding)
{
    
}

void xml_document::add_namespace(const std::string &id, const std::string &uri)
{
    
}

xml_node &xml_document::root()
{
    return root_;
}

const xml_node &xml_document::root() const
{
    return root_;
}

} // namespace xlnt