#pragma once

#include <string>
#include <vector>

namespace xlnt {
    
class worksheet;
    
std::vector<std::pair<std::string, std::string>> split_named_range(const std::string &named_range_string);
    
class named_range
{
public:
    named_range(const std::string &name, const std::vector<std::pair<worksheet, std::string>> &targets);
};
    
}