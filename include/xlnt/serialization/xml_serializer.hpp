#pragma once

#include <xlnt/utils/string.hpp>

#include <xlnt/xlnt_config.hpp>

namespace xlnt {

class xml_document;
class xml_node;

class XLNT_CLASS xml_serializer
{
  public:
    static string serialize(const xml_document &xml);
    static xml_document deserialize(const string &xml_string);

    static string serialize_node(const xml_node &xml);
};

} // namespace xlnt
