#pragma once

#include <string>
#include <vector>

namespace xlnt {

class relationship;
class zip_file;

class relationship_serializer
{
  public:
    static std::vector<relationship> read_relationships(zip_file &archive, const std::string &target);
    static bool write_relationships(const std::vector<relationship> &relationships, const std::string &target,
                                    zip_file &archive);
};

} // namespace xlnt
