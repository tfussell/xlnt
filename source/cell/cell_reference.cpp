#include <locale>

#include <xlnt/cell/cell_reference.hpp>
#include <xlnt/common/exceptions.hpp>
#include <xlnt/worksheet/range_reference.hpp>

#include "detail/constants.hpp"

namespace xlnt {
    
std::size_t cell_reference_hash::operator()(const cell_reference &k) const
{
    return k.get_row() * constants::MaxColumn + k.get_column_index();
}
    
cell_reference cell_reference::make_absolute(const cell_reference &relative_reference)
{
    cell_reference copy = relative_reference;
    copy.absolute_ = true;
    return copy;
}
    
cell_reference::cell_reference() : cell_reference(1, 1, false)
{
}

cell_reference::cell_reference(const std::string &string)
{
    bool absolute = false;
    auto split = split_reference(string, absolute, absolute);
    *this = cell_reference(split.first, split.second, absolute);
}

cell_reference::cell_reference(const char *reference_string) : cell_reference(std::string(reference_string))
{
    
}

cell_reference::cell_reference(const std::string &column, row_t row, bool absolute)
    : column_(column_index_from_string(column)),
      row_(row),
      absolute_(absolute)
{
    if(row == 0 || row_ >= constants::MaxRow || column_ == 0 || column_ >= constants::MaxColumn)
    {
        throw cell_coordinates_exception(column_, row_);
    }
}

cell_reference::cell_reference(column_t column_index, row_t row, bool absolute)
    : column_(column_index),
      row_(row),
      absolute_(absolute)
{
    if(row_ == 0 || row_ >= constants::MaxRow || column_ == 0 || column_ >= constants::MaxColumn)
    {
        throw cell_coordinates_exception(column_, row_);
    }
}

range_reference cell_reference::operator,(const xlnt::cell_reference &other) const
{
    return range_reference(*this, other);
}

std::string cell_reference::to_string() const
{
    if(absolute_)
    {
        return std::string("$") + column_string_from_index(column_) + "$" + std::to_string(row_);
    }
    
    return column_string_from_index(column_) + std::to_string(row_);
}

range_reference cell_reference::to_range() const
{
    return range_reference(column_, row_, column_, row_);
}

std::pair<std::string, row_t> cell_reference::split_reference(const std::string &reference_string, bool &absolute_column, bool &absolute_row)
{
    absolute_column = false;
    absolute_row = false;
    
    // Convert a coordinate string like 'B12' to a tuple ('B', 12)
    bool column_part = true;
    
    std::string column_string;
    
    for(auto character : reference_string)
    {
        char upper = std::toupper(character, std::locale::classic());
        
        if(std::isalpha(character, std::locale::classic()))
        {
            if(column_part)
            {
                column_string.append(1, upper);
            }
            else
            {
                throw cell_coordinates_exception(reference_string);
            }
        }
        else if(character == '$')
        {
            if(column_part)
            {
                if(column_string.empty())
                {
                    column_string.append(1, upper);
                }
                else
                {
                    column_part = false;
                }
            }
        }
        else
        {
            if(column_part)
            {
                column_part = false;
            }
            else if(!std::isdigit(character, std::locale::classic()))
            {
                throw cell_coordinates_exception(reference_string);
            }
        }
    }
    
    std::string row_string = reference_string.substr(column_string.length());
    
    if(row_string.length() == 0)
    {
        throw cell_coordinates_exception(reference_string);
    }

    if(column_string[0] == '$')
    {
        absolute_row = true;
        column_string = column_string.substr(1);
    }
    
    if(row_string[0] == '$')
    {
        absolute_column = true;
        row_string = row_string.substr(1);
    }
    
    return {column_string, std::stoi(row_string)};
}

cell_reference cell_reference::make_offset(int column_offset, int row_offset) const
{
    return cell_reference(static_cast<column_t>(static_cast<int>(column_) + column_offset),
        static_cast<row_t>(static_cast<int>(row_) + row_offset));
}

bool cell_reference::operator==(const cell_reference &comparand) const
{
    return comparand.column_ == column_
    && comparand.row_ == row_
    && absolute_ == comparand.absolute_;
}

column_t cell_reference::column_index_from_string(const std::string &column_string)
{
    if(column_string.length() > 3 || column_string.empty())
    {
        throw column_string_index_exception();
    }
    
    column_t column_index = 0;
    int place = 1;
    
    for(int i = static_cast<int>(column_string.length()) - 1; i >= 0; i--)
    {
        if(!std::isalpha(column_string[static_cast<std::size_t>(i)], std::locale::classic()))
        {
            throw column_string_index_exception();
        }
        
        column_index += static_cast<column_t>((std::toupper(column_string[static_cast<std::size_t>(i)], std::locale::classic()) - 'A' + 1) * place);
        place *= 26;
    }
    
    return column_index;
}

// Convert a column number into a column letter (3 -> 'C')
// Right shift the column col_idx by 26 to find column letters in reverse
// order.These numbers are 1 - based, and can be converted to ASCII
// ordinals by adding 64.
std::string cell_reference::column_string_from_index(column_t column_index)
{
    // these indicies corrospond to A->ZZZ and include all allowed
    // columns
    if(column_index < 1 || column_index > constants::MaxColumn)
    {
      //        auto msg = "Column index out of bounds: " + std::to_string(column_index);
        throw column_string_index_exception();
    }
    
    int temp = static_cast<int>(column_index);
    std::string column_letter = "";
    
    while(temp > 0)
    {
        int quotient = temp / 26, remainder = temp % 26;
        
        // check for exact division and borrow if needed
        if(remainder == 0)
        {
            quotient -= 1;
            remainder = 26;
        }
        
        column_letter = std::string(1, char(remainder + 64)) + column_letter;
        temp = quotient;
    }
    
    return column_letter;
}

bool cell_reference::operator<(const cell_reference &other)
{
    if(row_ != other.row_)
    {
        return row_ < other.row_;
    }

    return column_ < other.column_;
}

bool cell_reference::operator>(const cell_reference &other)
{
    if(row_ != other.row_)
    {
        return row_ > other.row_;
    }
    
    return column_ > other.column_;
}

bool cell_reference::operator<=(const cell_reference &other)
{
    if(row_ != other.row_)
    {
        return row_ < other.row_;
    }
    
    return column_ <= other.column_;
}

bool cell_reference::operator>=(const cell_reference &other)
{
    if(row_ != other.row_)
    {
        return row_ > other.row_;
    }
    
    return column_ >= other.column_;
}
    
} // namespace xlnt
