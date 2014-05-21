#include <locale>

#include "reference.h"
#include "custom_exceptions.h"

namespace xlnt {
    
std::size_t cell_reference_hash::operator()(const cell_reference &k) const
{
    return std::hash<std::string>()(k.to_string());
}
    
cell_reference cell_reference::make_absolute(const cell_reference &relative_reference)
{
    cell_reference copy = relative_reference;
    copy.absolute_ = true;
    return copy;
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

cell_reference::cell_reference(const std::string &column, int row, bool absolute)
: column_index_(column_index_from_string(column) - 1),
row_index_(row - 1),
absolute_(absolute)
{
    if(row_index_ < 0 || column_index_ < 0 || row_index_ >= 1048576 || column_index_ >= 16384)
    {
        throw bad_cell_coordinates(column_index_, row_index_);
    }
}

cell_reference::cell_reference(int column_index, int row_index, bool absolute)
: column_index_(column_index),
row_index_(row_index),
absolute_(absolute)
{
    if(row_index_ < 0 || column_index_ < 0 || row_index_ >= 1048576 || column_index_ >= 16384)
    {
        throw bad_cell_coordinates(column_index_, row_index_);
    }
}

std::string cell_reference::to_string() const
{
    if(absolute_)
    {
        return std::string("$") + column_string_from_index(column_index_ + 1) + "$" + std::to_string(row_index_ + 1);
    }
    return column_string_from_index(column_index_ + 1) + std::to_string(row_index_ + 1);
}

std::pair<std::string, int> cell_reference::split_reference(const std::string &reference_string, bool &absolute_column, bool &absolute_row)
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
                throw bad_cell_coordinates(reference_string);
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
                throw bad_cell_coordinates(reference_string);
            }
        }
    }
    
    std::string row_string = reference_string.substr(column_string.length());
    
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

int cell_reference::column_index_from_string(const std::string &column_string)
{
    if(column_string.length() > 3 || column_string.empty())
    {
        throw std::runtime_error("column must be one to three characters");
    }
    
    int column_index = 0;
    int place = 1;
    
    for(int i = static_cast<int>(column_string.length()) - 1; i >= 0; i--)
    {
        if(!std::isalpha(column_string[i], std::locale::classic()))
        {
            throw std::runtime_error("column must contain only letters in the range A-Z");
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
std::string cell_reference::column_string_from_index(int column_index)
{
    // these indicies corrospond to A->ZZZ and include all allowed
    // columns
    if(column_index < 1 || column_index > 18278)
    {
        auto msg = "Column index out of bounds: " + std::to_string(column_index);
        throw std::runtime_error(msg);
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

range_reference range_reference::make_absolute(const xlnt::range_reference &relative_reference)
{
    range_reference copy = relative_reference;
    copy.top_left_.set_absolute(true);
    copy.bottom_right_.set_absolute(true);
    return copy;
}

range_reference::range_reference() : range_reference("A1")
{
}

range_reference::range_reference(const char *range_string) : range_reference(std::string(range_string))
{
}

range_reference::range_reference(const std::string &range_string)
: top_left_("A1"),
bottom_right_("A1")
{
    auto colon_index = range_string.find(':');
    
    if(colon_index != std::string::npos)
    {
        top_left_ = cell_reference(range_string.substr(0, colon_index));
        bottom_right_ = cell_reference(range_string.substr(colon_index + 1));
    }
    else
    {
        top_left_ = cell_reference(range_string);
        bottom_right_ = cell_reference(range_string);
    }
}

range_reference::range_reference(const cell_reference &top_left, const cell_reference &bottom_right)
: top_left_(top_left),
bottom_right_(bottom_right)
{
    
}

range_reference::range_reference(int column_index_start, int row_index_start, int column_index_end, int row_index_end)
: top_left_(column_index_start, row_index_start),
bottom_right_(column_index_end, row_index_end)
{
    
}

range_reference range_reference::make_offset(int column_offset, int row_offset) const
{
    return range_reference(top_left_.make_offset(column_offset, row_offset), bottom_right_.make_offset(column_offset, row_offset));
}

int range_reference::get_height() const
{
    return bottom_right_.get_row_index() - top_left_.get_row_index();
}

int range_reference::get_width() const
{
    return bottom_right_.get_column_index() - top_left_.get_column_index();
}

bool range_reference::is_single_cell() const
{
    return get_width() == 0 && get_height() == 0;
}

std::string range_reference::to_string() const
{
    if(is_single_cell())
    {
        return top_left_.to_string();
    }
    return top_left_.to_string() + ":" + bottom_right_.to_string();
}

bool range_reference::operator==(const range_reference &comparand) const
{
    return comparand.top_left_ == top_left_
    && comparand.bottom_right_ == bottom_right_;
}

    
}