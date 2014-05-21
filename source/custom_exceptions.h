#pragma once

#include <exception>
#include <string>

namespace xlnt {

class bad_sheet_title : public std::runtime_error
{
public:
    bad_sheet_title(const std::string &title)
    : std::runtime_error(std::string("bad worksheet title: ") + title)
    {
        
    }
};

class bad_cell_coordinates : public std::runtime_error
{
public:
    bad_cell_coordinates(int row, int column)
    : std::runtime_error(std::string("bad cell coordinates: (") + std::to_string(row) + "," + std::to_string(column) + ")")
    {
        
    }
    
    bad_cell_coordinates(const std::string &coord_string)
    : std::runtime_error(std::string("bad cell coordinates: (") + coord_string + ")")
    {
        
    }
};
    
} // namespace xlnt
