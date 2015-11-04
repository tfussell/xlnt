#include <sstream>

#include <xlnt/serialization/xml_serializer.hpp>
#include <xlnt/serialization/xml_document.hpp>
#include <xlnt/serialization/xml_node.hpp>

#include "detail/include_pugixml.hpp"
#include "detail/xml_document_impl.hpp"
#include "detail/xml_node_impl.hpp"

namespace xlnt {

string xml_serializer::serialize(const xml_document &xml)
{
    std::ostringstream ss;
    xml.d_->doc.save(ss, "   ", pugi::format_default, pugi::encoding_utf8);

    return ss.str().data();
}

xml_document xml_serializer::deserialize(const string &xml_string)
{
    xml_document doc;
    doc.d_->doc.load(xml_string.data());

    return doc;
}

string xml_serializer::serialize_node(const xml_node &xml)
{
    pugi::xml_document doc;
    doc.append_copy(xml.d_->node);

    std::ostringstream ss;
    doc.save(ss);

    return ss.str().data();
}

} // namespace xlnt
