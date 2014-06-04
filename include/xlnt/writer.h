#pragma once

#include <string>
#include <vector>

namespace xlnt {
    
class workbook;
class worksheet;
    
class writer
{
public:
    static std::string write_workbook(const workbook &wb);
    static std::string write_worksheet(worksheet ws);
    static std::string write_worksheet(worksheet ws, const std::vector<std::string> &string_table);
    static std::string write_worksheet(worksheet ws, const std::vector<std::string> &string_table, const std::unordered_map<std::size_t, std::string> &style_table);
    static std::string write_theme();
    static std::string write_content_types(const std::pair<std::unordered_map<std::string, std::string>, std::unordered_map<std::string, std::string>> &content_types);
    static std::string write_relationships(const std::unordered_map<std::string, std::pair<std::string, std::string>> &relationships);
    static std::string write_workbook_rels(const workbook &wb);
    static std::string write_worksheet_rels(worksheet ws, int n);
    static std::string write_string_table(const std::unordered_map<std::string, int> &string_table);
};
    
} // namespace xlnt
