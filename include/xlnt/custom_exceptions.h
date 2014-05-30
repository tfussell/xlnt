#pragma once

#include <stdexcept>
#include <string>

namespace xlnt {

class bad_sheet_title : public std::runtime_error
{
public:
    bad_sheet_title(const std::string &title);
};

class cell_coordinates_exception : public std::runtime_error
{
public:
    cell_coordinates_exception(int row, int column);
    
    cell_coordinates_exception(const std::string &coord_string);
};

class workbook_already_saved : public std::runtime_error
{
public:
    workbook_already_saved();
};

class data_type_exception : public std::runtime_error
{
public:
    data_type_exception();
};

class column_string_index_exception : public std::runtime_error
{
public:
    column_string_index_exception();
};
    
} // namespace xlnt
