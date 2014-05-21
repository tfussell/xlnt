#pragma once

#include "cell_reference.h"

namespace xlnt {

class range_reference
{
public:
    /// <summary>
    /// Convert a coordinate to an absolute coordinate string (B12 -> $B$12)
    /// </summary>
    static range_reference make_absolute(const range_reference &relative_reference);

    range_reference();
    range_reference(const std::string &range_string);
    range_reference(const char *range_string);
    range_reference(const std::pair<cell_reference, cell_reference> &reference_pair);
    range_reference(const cell_reference &start, const cell_reference &end);
    range_reference(column_t column_index_start, row_t row_index_start, column_t column_index_end, row_t row_index_end);

    bool is_single_cell() const;
    column_t get_width() const;
    row_t get_height() const;
    cell_reference get_top_left() const { return top_left_; }
    cell_reference get_bottom_right() const { return bottom_right_; }
    cell_reference &get_top_left() { return top_left_; }
    cell_reference &get_bottom_right() { return bottom_right_; }

    range_reference make_offset(column_t column_offset, row_t row_offset) const;

    std::string to_string() const;

    bool operator==(const range_reference &comparand) const;
    bool operator==(const std::string &reference_string) const { return *this == range_reference(reference_string); }
    bool operator==(const char *reference_string) const { return *this == std::string(reference_string); }

private:
    cell_reference top_left_;
    cell_reference bottom_right_;
};

} // namespace xlnt
