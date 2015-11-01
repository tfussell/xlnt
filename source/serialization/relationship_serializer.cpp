#include <xlnt/serialization/relationship_serializer.hpp>

#include <xlnt/common/relationship.hpp>
#include <xlnt/common/zip_file.hpp>
#include <xlnt/serialization/xml_document.hpp>
#include <xlnt/serialization/xml_node.hpp>

#include "detail/constants.hpp"

namespace {

std::string make_rels_name(const std::string &target)
{
    const char sep = '/';

    if (target.empty() || target.back() == sep)
    {
        return target + "_rels/.rels";
    }

    auto sep_pos = target.find_last_of(sep);
    auto first_part = target.substr(0, sep_pos + 1);
    auto last_part = target.substr(sep_pos + 1);

    return first_part + "_rels/" + last_part + ".rels";
}
}

namespace xlnt {

std::vector<relationship> relationship_serializer::read_relationships(zip_file &archive, const std::string &target)
{
    xml_document xml;
    xml.from_string(archive.read(make_rels_name(target)));

    auto root_node = xml.get_child("Relationships");

    std::vector<relationship> relationships;

    for (auto relationship_node : root_node.get_children())
    {
        if (relationship_node.get_name() != "Relationship")
        {
            continue;
        }

        std::string id = relationship_node.get_attribute("Id");
        std::string type = relationship_node.get_attribute("Type");
        std::string target = relationship_node.get_attribute("Target");

        relationships.push_back(xlnt::relationship(type, id, target));
    }

    return relationships;
}

bool relationship_serializer::write_relationships(const std::vector<relationship> &relationships,
                                                  const std::string &target, zip_file &archive)
{
    xml_document xml;

    auto root_node = xml.add_child("Relationships");

    xml.add_namespace("", constants::Namespaces.at("relationships"));

    for (const auto &relationship : relationships)
    {
        auto target = relationship.get_target_uri();

        auto relationship_node = root_node.add_child("Relationship");

        relationship_node.add_attribute("Id", relationship.get_id());
        relationship_node.add_attribute("Target", target);
        relationship_node.add_attribute("Type", relationship.get_type_string());

        if (relationship.get_target_mode() == target_mode::external)
        {
            relationship_node.add_attribute("TargetMode", "External");
        }
    }

    archive.writestr(make_rels_name(target), xml.to_string());

    return true;
}
}
