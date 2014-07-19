#include "common/exceptions.hpp"

namespace xlnt {

sheet_title_exception::sheet_title_exception(const std::string &title)
    : std::runtime_error(std::string("bad worksheet title: ") + title)
{

}

column_string_index_exception::column_string_index_exception()
    : std::runtime_error("")
{

}

data_type_exception::data_type_exception()
    : std::runtime_error("")
{

}

invalid_file_exception::invalid_file_exception(const std::string &filename)
    : std::runtime_error(std::string("couldn't open file: (") + filename + ")")
{
    
}

cell_coordinates_exception::cell_coordinates_exception(int row, int column)
    : std::runtime_error(std::string("bad cell coordinates: (") + std::to_string(row) + "," + std::to_string(column) + ")")
{

}

cell_coordinates_exception::cell_coordinates_exception(const std::string &coord_string)
    : std::runtime_error(std::string("bad cell coordinates: (") + coord_string + ")")
{

}

} // namespace xlnt
