#pragma once

#include <stdexcept>
#include <string>

namespace xlnt {

class bad_sheet_title : public std::runtime_error
{
public:
    bad_sheet_title(const std::string &title);
};

class bad_cell_coordinates : public std::runtime_error
{
public:
    bad_cell_coordinates(int row, int column);
    
    bad_cell_coordinates(const std::string &coord_string);
};
    
} // namespace xlnt
