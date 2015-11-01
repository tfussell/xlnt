#pragma once

#include <string>

namespace xlnt {

class xml_document;
class xml_node;

class xml_serializer
{
  public:
    static std::string serialize(const xml_document &xml);
    static xml_document deserialize(const std::string &xml_string);

    static std::string serialize_node(const xml_node &xml);
};

} // namespace xlnt
