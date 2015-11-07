#include <xlnt/utils/exceptions.hpp>

namespace xlnt {

sheet_title_exception::sheet_title_exception(const string &title)
{
}

column_string_index_exception::column_string_index_exception()
{
}

data_type_exception::data_type_exception()
{
}

attribute_error::attribute_error()
{
}

named_range_exception::named_range_exception()
{
}

invalid_file_exception::invalid_file_exception(const string &filename)
{
}

cell_coordinates_exception::cell_coordinates_exception(column_t column, row_t row)
{
}

cell_coordinates_exception::cell_coordinates_exception(const string &coord_string)
{
}

illegal_character_error::illegal_character_error(xlnt::utf32_char c)
{
}

} // namespace xlnt
