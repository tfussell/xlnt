#pragma once

#include <string>
#include <vector>

namespace xlnt {
    
class relationship;
class xml_document;

class relationship_serializer
{
public:
    static bool read_relationships(const xml_document &xml, const std::string &dir, std::vector<relationship> &relationships);
    static bool write_relationships(const std::vector<relationship> &relationships, const std::string &dir, xml_document &xml);
};
    
} // namespace xlnt
