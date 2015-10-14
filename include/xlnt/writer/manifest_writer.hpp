#pragma once

#include <string>

namespace xlnt {
    
class workbook;

std::string write_content_types(const workbook &wb, bool as_template);
    
};
