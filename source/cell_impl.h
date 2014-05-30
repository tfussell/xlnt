#pragma once

#include "types.h"
#include "relationship.h"

namespace xlnt {
    
class style;
    
struct cell_impl
{
    cell_impl();
    cell_impl(int column_index, int row_index);
    
    std::string to_string() const;
    
    enum class type
    {
        null,
        numeric,
        string,
        date,
        formula,
        boolean,
        error,
        hyperlink
    };
    
    type type_;
    long double numeric_value;
    std::string error_value;
    tm date_value;
    std::string string_value;
    std::string formula_value;
    column_t column;
    row_t row;
    style *style_;
    relationship hyperlink_rel;
    bool merged;
};
    
}