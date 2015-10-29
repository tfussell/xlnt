#include <xlnt/s11n/relationship_serializer.hpp>

#include <xlnt/common/relationship.hpp>
#include <xlnt/s11n/xml_document.hpp>
#include <xlnt/s11n/xml_node.hpp>

#include "detail/constants.hpp"

namespace xlnt {

bool relationship_serializer::read_relationships(const xml_document &xml, const std::string &dir, std::vector<relationship> &relationships)
{
    return false;
}
    
bool relationship_serializer::write_relationships(const std::vector<relationship> &relationships, const std::string &dir, xml_document &xml)
{
    xml.add_namespace("", constants::Namespaces.at("relationships"));
    
    auto root_node = xml.root();
    root_node.set_name("Relationships");
    
    for(const auto &relationship : relationships)
    {
        auto target = relationship.get_target_uri();
        
        if (dir != "" && target.substr(0, dir.size()) == dir)
        {
            target = target.substr(dir.size());
        }
        
        auto app_props_node = root_node.add_child("Relationship");
        
        app_props_node.add_attribute("Id", relationship.get_id());
        app_props_node.add_attribute("Target", target);
        app_props_node.add_attribute("Type", relationship.get_type_string());
        
        if(relationship.get_target_mode() == target_mode::external)
        {
            app_props_node.add_attribute("TargetMode", "External");
        }
    }

    return true;
}
    
}
