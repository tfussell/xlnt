#include <sstream>

#include <xlnt/s11n/xml_serializer.hpp>
#include <xlnt/s11n/xml_document.hpp>
#include <xlnt/s11n/xml_node.hpp>

#include "detail/include_pugixml.hpp"
#include "detail/xml_document_impl.hpp"
#include "detail/xml_node_impl.hpp"

namespace xlnt {

std::string xml_serializer::serialize(const xml_document &xml)
{
    std::ostringstream ss;
    xml.d_->doc.save(ss, "   ", pugi::format_default, pugi::encoding_utf8);

    return ss.str();
}

xml_document xml_serializer::deserialize(const std::string &xml_string)
{
    xml_document doc;
    doc.d_->doc.load(xml_string.c_str());

    return doc;
}

std::string xml_serializer::serialize_node(const xml_node &xml)
{
    pugi::xml_document doc;
    doc.append_copy(xml.d_->node);

    std::ostringstream ss;
    doc.save(ss);

    return ss.str();
}

} // namespace xlnt
