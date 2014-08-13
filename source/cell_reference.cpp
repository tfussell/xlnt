#include <locale>

#include <xlnt/cell/cell_reference.hpp>
#include <xlnt/common/exceptions.hpp>
#include <xlnt/worksheet/range_reference.hpp>

#include "constants.hpp"

namespace xlnt {
    
std::size_t cell_reference_hash::operator()(const cell_reference &k) const
{
    return k.get_row_index() * constants::MaxColumn + k.get_column_index();
}
    
cell_reference cell_reference::make_absolute(const cell_reference &relative_reference)
{
    cell_reference copy = relative_reference;
    copy.absolute_ = true;
    return copy;
}
    
cell_reference::cell_reference() : cell_reference(0, 0, false)
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
: column_index_(column_index_from_string(column) - 1),
row_index_(row - 1),
absolute_(absolute)
{
    if(row == 0 || row_index_ >= constants::MaxRow || column_index_ >= constants::MaxColumn)
    {
        throw cell_coordinates_exception(column_index_, row_index_);
    }
}

cell_reference::cell_reference(column_t column_index, row_t row_index, bool absolute)
: column_index_(column_index),
row_index_(row_index),
absolute_(absolute)
{
    if(row_index_ >= constants::MaxRow || column_index_ >= constants::MaxColumn)
    {
        throw cell_coordinates_exception(column_index_, row_index_);
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
        return std::string("$") + column_string_from_index(column_index_ + 1) + "$" + std::to_string(row_index_ + 1);
    }
    return column_string_from_index(column_index_ + 1) + std::to_string(row_index_ + 1);
}

range_reference cell_reference::to_range() const
{
    return range_reference(column_index_, row_index_, column_index_, row_index_);
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
        else
        {
            if(column_part)
            {
                column_part = false;
            }
            else if(!(std::isdigit(character, std::locale::classic()) || character == '$'))
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
    return cell_reference(column_index_ + column_offset, row_index_ + row_offset);
}

bool cell_reference::operator==(const cell_reference &comparand) const
{
    return comparand.column_index_ == column_index_
    && comparand.row_index_ == row_index_
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
        if(!std::isalpha(column_string[i], std::locale::classic()))
        {
	    throw column_string_index_exception();
        }
        
        column_index += (std::toupper(column_string[i], std::locale::classic()) - 'A' + 1) * place;
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
    
    auto temp = column_index;
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

bool operator<(const cell_reference &left, const cell_reference &right)
{
    if(left.row_index_ != right.row_index_)
    {
        return left.row_index_ < right.row_index_;
    }

    return left.column_index_ < right.column_index_;
}
    
}
