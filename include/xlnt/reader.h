#pragma once

#include <iostream>
#include <string>
#include <unordered_map>
#include <vector>

namespace xlnt {
    
class workbook;
class worksheet;
    
class reader
{
public:
    static std::unordered_map<std::string, std::pair<std::string, std::string>> read_relationships(const std::string &content);
    static std::pair<std::unordered_map<std::string, std::string>, std::unordered_map<std::string, std::string>> read_content_types(const std::string &content);
    static std::string determine_document_type(const std::unordered_map<std::string, std::pair<std::string, std::string>> &root_relationships,
                                               const std::unordered_map<std::string, std::string> &override_types);
    static worksheet read_worksheet(std::istream &handle, workbook &wb, const std::string &title, const std::vector<std::string> &string_table);
    static void read_worksheet(worksheet ws, const std::string &xml_string, const std::vector<std::string> &string_table);
    static std::vector<std::string> read_shared_string(const std::string &xml_string);
};
    
} // namespace xlnt
