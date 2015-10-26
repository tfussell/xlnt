#pragma once

#include <string>
#include <unordered_map>
#include <vector>

namespace xlnt {
    
class worksheet;

std::string write_worksheet(worksheet ws, const std::vector<std::string> &string_table);
    
} // namespace xlnt
