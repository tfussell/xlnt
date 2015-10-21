#include <sstream>

#include <xlnt/workbook/workbook.hpp>
#include <xlnt/writer/manifest_writer.hpp>

#include "constants.hpp"
#include "detail/include_pugixml.hpp"

namespace xlnt {

std::string write_content_types(const workbook &wb, bool /*as_template*/)
{
    pugi::xml_document doc;
    auto root_node = doc.append_child("Types");
    root_node.append_attribute("xmlns").set_value(constants::Namespaces.at("content-types").c_str());
    
    for(auto type : wb.get_content_types())
    {
        pugi::xml_node type_node;
        
        if (type.is_default)
        {
            type_node = root_node.append_child("Default");
            type_node.append_attribute("Extension").set_value(type.extension.c_str());
        }
        else
        {
            type_node = root_node.append_child("Override");
            type_node.append_attribute("PartName").set_value(type.part_name.c_str());
        }
        
        type_node.append_attribute("ContentType").set_value(type.type.c_str());
    }
    
    std::stringstream ss;
    doc.save(ss);
    return ss.str();
}

} // namespace xlnt
