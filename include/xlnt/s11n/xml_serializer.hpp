#pragma once

#include <string>

namespace xlnt {
    
class xml_document;

class xml_serializer
{
public:
    static std::string serialize(const xml_document &xml);
    static xml_document deserialize(const std::string &xml_string);
};

} // namespace xlnt
