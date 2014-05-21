#pragma once

#include <string>
#include <utility>

namespace xlnt {
    
class cell_reference;
    
struct cell_reference_hash
{
    std::size_t operator()(const cell_reference &k) const;
};
    
class cell_reference
{
public:
    /// <summary>
    /// Convert a coordinate to an absolute coordinate string (B12 -> $B$12)
    /// </summary>
    static cell_reference make_absolute(const cell_reference &relative_reference);
    
    /// <summary>
    /// Convert a column letter into a column number (e.g. B -> 2)
    /// </summary>
    /// <remarks>
    /// Excel only supports 1 - 3 letter column names from A->ZZZ, so we
    /// restrict our column names to 1 - 3 characters, each in the range A - Z.
    /// </remarks>
    static int column_index_from_string(const std::string &column_string);
    
    /// <summary>
    /// Convert a column number into a column letter (3 -> 'C')
    /// </summary>
    /// <remarks>
    /// Right shift the column col_idx by 26 to find column letters in reverse
    /// order.These numbers are 1 - based, and can be converted to ASCII
    /// ordinals by adding 64.
    /// </remarks>
    static std::string column_string_from_index(int column_index);
    
    static std::pair<std::string, int> split_reference(const std::string &reference_string,
        bool &absolute_column, bool &absolute_row);
    
    cell_reference(const char *reference_string);
    cell_reference(const std::string &reference_string);
    cell_reference(const std::string &column, int row, bool absolute = false);
    cell_reference(int column, int row, bool absolute = false);
    
    bool is_absolute() const { return absolute_; }
    void set_absolute(bool absolute) { absolute_ = absolute; }
    
    std::string get_column() const { return column_string_from_index(column_index_ + 1); }
    void set_column(const std::string &column) { column_index_ = column_index_from_string(column) - 1; }
    
    int get_column_index() const { return column_index_; }
    void set_column_index(int column_index) { column_index_ = column_index; }
    
    int get_row() const { return row_index_ + 1; }
    void set_row(int row) { row_index_ = row - 1; }
    
    int get_row_index() const { return row_index_; }
    void set_row_index(int row_index) { row_index_ = row_index; }
    
    cell_reference make_offset(int column_offset, int row_offset) const;
    
    std::string to_string() const;
    
    bool operator==(const cell_reference &comparand) const;
    bool operator==(const std::string &reference_string) const { return *this == cell_reference(reference_string); }
    bool operator==(const char *reference_string) const { return *this == std::string(reference_string); }
    
private:
    int column_index_;
    int row_index_;
    bool absolute_;
};
    
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
    range_reference(const std::string &column_start, int row_start, const std::string &column_end, int row_end);
    range_reference(int column_start, int row_start, int column_end, int row_end);
    
    bool is_single_cell() const;
    int get_width() const;
    int get_height() const;
    cell_reference get_top_left() const { return top_left_; }
    cell_reference get_bottom_right() const { return bottom_right_; }
    cell_reference &get_top_left() { return top_left_; }
    cell_reference &get_bottom_right() { return bottom_right_; }
    
    range_reference make_offset(int column_offset, int row_offset) const;
    
    std::string to_string() const;
    
    bool operator==(const range_reference &comparand) const;
    bool operator==(const std::string &reference_string) const { return *this == range_reference(reference_string); }
    bool operator==(const char *reference_string) const { return *this == std::string(reference_string); }
    
private:
    cell_reference top_left_;
    cell_reference bottom_right_;
};
    
} // namespace xlnt
