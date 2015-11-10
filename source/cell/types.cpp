#include <locale>

#include <xlnt/cell/types.hpp>
#include <xlnt/utils/exceptions.hpp>

#include <detail/constants.hpp>

namespace xlnt {

column_t::index_t column_t::column_index_from_string(const std::string &column_string)
{
    if (column_string.length() > 3 || column_string.empty())
    {
        throw column_string_index_exception();
    }

    column_t::index_t column_index = 0;
    int place = 1;

    for (int i = static_cast<int>(column_string.length()) - 1; i >= 0; i--)
    {
        if (!std::isalpha(column_string[static_cast<std::size_t>(i)], std::locale::classic()))
        {
            throw column_string_index_exception();
        }

        auto char_index = std::toupper(column_string[static_cast<std::size_t>(i)], std::locale::classic()) - 'A';

        column_index += static_cast<column_t::index_t>((char_index + 1) * place);
        place *= 26;
    }

    return column_index;
}

// Convert a column number into a column letter (3 -> 'C')
// Right shift the column col_idx by 26 to find column letters in reverse
// order.These numbers are 1 - based, and can be converted to ASCII
// ordinals by adding 64.
std::string column_t::column_string_from_index(column_t::index_t column_index)
{
    // these indicies corrospond to A->ZZZ and include all allowed
    // columns
    if (column_index < constants::MinColumn() || column_index > constants::MaxColumn())
    {
        //        auto msg = "Column index out of bounds: " + std::to_string(column_index);
        throw column_string_index_exception();
    }

    int temp = static_cast<int>(column_index);
    std::string column_letter = "";

    while (temp > 0)
    {
        int quotient = temp / 26, remainder = temp % 26;

        // check for exact division and borrow if needed
        if (remainder == 0)
        {
            quotient -= 1;
            remainder = 26;
        }

        column_letter = std::string(1, char(remainder + 64)) + column_letter;
        temp = quotient;
    }

    return column_letter;
}

std::size_t column_hash::operator()(const column_t &k) const
{
    static std::hash<column_t::index_t> hasher;
    return hasher(k.index);
}

} // namespace xlnt
