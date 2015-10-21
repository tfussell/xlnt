#include <sstream>

#include <xlnt/common/relationship.hpp>
#include <xlnt/writer/relationship_writer.hpp>

#include "constants.hpp"
#include "detail/include_pugixml.hpp"

namespace xlnt {

std::string write_relationships(const std::vector<relationship> &relationships, const std::string &dir)
{
    pugi::xml_document doc;
    
    auto root_node = doc.append_child("Relationships");
    root_node.append_attribute("xmlns").set_value(constants::Namespaces.at("relationships").c_str());
    
    for(auto relationship : relationships)
    {
        auto target = relationship.get_target_uri();
        
        if (dir != "" && target.substr(0, dir.size()) == dir)
        {
            target = target.substr(dir.size());
        }
        
        auto app_props_node = root_node.append_child("Relationship");
        
        app_props_node.append_attribute("Id").set_value(relationship.get_id().c_str());
        app_props_node.append_attribute("Target").set_value(target.c_str());
        app_props_node.append_attribute("Type").set_value(relationship.get_type_string().c_str());
        
        if(relationship.get_target_mode() == target_mode::external)
        {
            app_props_node.append_attribute("TargetMode").set_value("External");
        }
    }
    
    std::stringstream ss;
    doc.save(ss);
    
    return ss.str();
}
    
}
