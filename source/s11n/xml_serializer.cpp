#include <sstream>

#include <xlnt/s11n/xml_serializer.hpp>
#include <xlnt/s11n/xml_document.hpp>
#include <xlnt/s11n/xml_node.hpp>

#include "detail/include_pugixml.hpp"

namespace {

void serialize_node(const xlnt::xml_node &source, pugi::xml_node node)
{
    for(const auto &attribute : source.get_attributes())
    {
        node.append_attribute(attribute.first.c_str()).set_value(attribute.second.c_str());
    }
    
    if(source.has_text())
    {
        node.text().set(source.get_text().c_str());
    }
    else
    {
        for(const auto &child : source.get_children())
        {
            pugi::xml_node child_node = node.append_child(child.get_name().c_str());
            serialize_node(child, child_node);
        }
    }
}

xlnt::xml_node deserialize_node(const pugi::xml_node source)
{
    xlnt::xml_node node(source.name());
    
    for(const auto attribute : source.attributes())
    {
        node.add_attribute(attribute.name(), attribute.as_string());
    }

    if(source.text() != nullptr)
    {
        node.set_text(source.text().as_string());
    }
    else
    {
        for(const auto child : source.children())
        {
            node.add_child(deserialize_node(child));
        }
    }
    
    return node;
}
    
} // namespace

namespace xlnt {

std::string xml_serializer::serialize(const xml_document &xml)
{
    pugi::xml_document doc;
    
    auto root = doc.root();
    root.set_name(xml.root().get_name().c_str());
    serialize_node(xml.root(), root);
    
    std::ostringstream ss;
    doc.save(ss);
    
    return ss.str();
}
    
xml_document xml_serializer::deserialize(const std::string &xml_string)
{
    pugi::xml_document doc;
    doc.load(xml_string.c_str());
    
    xml_document result;
    result.root() = deserialize_node(doc.root());
    
    return result;
}
    
} // namespace xlnt
