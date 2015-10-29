#include <xlnt/s11n/manifest_serializer.hpp>
#include <xlnt/s11n/xml_document.hpp>
#include <xlnt/workbook/manifest.hpp>

#include "detail/constants.hpp"

namespace xlnt {

bool manifest_serializer::write_manifest(xml_document &xml)
{
    xml.add_namespace("", constants::Namespaces.at("content-types"));
    
    auto root_node = xml.root();
    root_node.set_name("Types");
    
    for(const auto &default_type : manifest_.get_default_types())
    {
        auto type_node = root_node.add_child("Default");
        type_node.add_attribute("Extension", default_type.get_extension());
        type_node.add_attribute("ContentType", default_type.get_extension());
    }
    
    for(const auto &override_type : manifest_.get_override_types())
    {
        auto type_node = root_node.add_child("Override");
        type_node.add_attribute("PartName", override_type.get_part_name());
        type_node.add_attribute("ContentType", override_type.get_content_type());
    }
    
    return true;
}

} // namespace xlnt
