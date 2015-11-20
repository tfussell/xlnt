#include <xlnt/utils/exceptions.hpp>

namespace xlnt {

sheet_title_exception::sheet_title_exception(const std::string &title)
    : std::runtime_error(std::string("bad worksheet title: ") + title)
{
}

column_string_index_exception::column_string_index_exception()
    : std::runtime_error("column string index exception")
{
}

data_type_exception::data_type_exception()
    : std::runtime_error("data type exception")
{
}

attribute_error::attribute_error()
    : std::runtime_error("attribute error")
{
}

named_range_exception::named_range_exception()
    : std::runtime_error("named range not found or not owned by this worksheet")
{
}

invalid_file_exception::invalid_file_exception(const std::string &filename)
    : std::runtime_error(std::string("couldn't open file: (") + filename + ")")
{
}

cell_coordinates_exception::cell_coordinates_exception(column_t column, row_t row)
    : std::runtime_error(std::string("bad cell coordinates: (") + std::to_string(column.index) + ", " + std::to_string(row) +
                         ")")
{
}

cell_coordinates_exception::cell_coordinates_exception(const std::string &coord_string)
    : std::runtime_error(std::string("bad cell coordinates: (") + (coord_string.empty() ? "<empty>" : coord_string) + ")")
{
}

illegal_character_error::illegal_character_error(char c)
    : std::runtime_error(std::string("illegal character: (") + std::to_string(static_cast<unsigned char>(c)) + ")")
{
}

value_error::value_error()
    : std::runtime_error("value error")
{
}

} // namespace xlnt
