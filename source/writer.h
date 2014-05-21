#pragma once

#include <string>
#include <vector>

namespace xlnt {
    
class worksheet;
    
class writer
{
public:
    static std::string write_worksheet(worksheet ws);
    static std::string write_worksheet(worksheet ws, const std::vector<std::string> &string_table);
    static std::string write_theme();
    static std::string write_content_types(const std::pair<std::unordered_map<std::string, std::string>, std::unordered_map<std::string, std::string>> &content_types);
    static std::string write_relationships(const std::unordered_map<std::string, std::pair<std::string, std::string>> &relationships);
};
    
} // namespace xlnt
