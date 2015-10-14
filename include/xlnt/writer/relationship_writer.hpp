#pragma once

#include <string>
#include <vector>

namespace xlnt {
    
class relationship;
 
std::string write_relationships(const std::vector<relationship> &relationships, const std::string &dir);
    
} // namespace xlnt
