#include <locale>

#include "worksheet/range_reference.hpp"

namespace xlnt {

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

range_reference::range_reference(column_t column_index_start, row_t row_index_start, column_t column_index_end, row_t row_index_end)
    : top_left_(column_index_start, row_index_start),
    bottom_right_(column_index_end, row_index_end)
{

}

range_reference range_reference::make_offset(column_t column_offset, row_t row_offset) const
{
    return range_reference(top_left_.make_offset(column_offset, row_offset), bottom_right_.make_offset(column_offset, row_offset));
}

row_t range_reference::get_height() const
{
    return bottom_right_.get_row_index() - top_left_.get_row_index();
}

column_t range_reference::get_width() const
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
