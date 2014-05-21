#pragma once

#include "cell.h"
#include "relationship.h"

namespace xlnt {

struct cell_struct
{
    cell_struct(worksheet_struct *ws, int column, int row)
        : type(cell::type::null), parent_worksheet(ws),
        column(column), row(row),
        hyperlink_rel("invalid", "")
    {

    }

    cell_struct()
        : type(cell::type::null), parent_worksheet(nullptr),
        column(0), row(0), hyperlink_rel("invalid", "")
    {

    }

    std::string to_string() const;

    cell::type type;

    union
    {
        long double numeric_value;
        bool bool_value;
    };

    std::string error_value;
    tm date_value;
    std::string string_value;
    std::string formula_value;
    worksheet_struct *parent_worksheet;
    column_t column;
    row_t row;
    style *style;
    relationship hyperlink_rel;
    bool merged;
};

} // namespace xlnt
