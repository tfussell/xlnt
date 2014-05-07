#include "worksheet.h"

namespace xlnt {

namespace {

class not_implemented : public std::runtime_error
{
public:
    not_implemented() : std::runtime_error("error: not implemented")
    {

    }
};

// Convert a column number into a column letter (3 -> 'C')
// Right shift the column col_idx by 26 to find column letters in reverse
// order.These numbers are 1 - based, and can be converted to ASCII
// ordinals by adding 64.
std::string get_column_letter(int column_index)
{
    // these indicies corrospond to A->ZZZ and include all allowed
    // columns
    if(column_index < 1 || column_index > 18278)
    {
        auto msg = "Column index out of bounds: " + column_index;
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

}

class worksheet_impl
{
public:
    worksheet_impl(workbook &parent_workbook, const std::string &title) 
        : parent_(parent_workbook), title_(title)
    {
        
    }

    void garbage_collect()
    {
        throw not_implemented();
    }

    std::set<cell> get_cell_collection()
    {
        throw not_implemented();
    }

    std::string get_title() const
    {
        throw not_implemented();
    }

    void set_title(const std::string &title)
    {
        title_ = title;
    }

    cell get_freeze_panes() const
    {
        throw not_implemented();
    }

    void set_freeze_panes(cell top_left_cell)
    {
        throw not_implemented();
    }

    void set_freeze_panes(const std::string &top_left_coordinate)
    {
        freeze_panes_ = cell(top_left_coordinate);
    }

    void unfreeze_panes()
    {
        throw not_implemented();
    }

    xlnt::cell cell(const std::string &coordinate)
    {
        return cells_[coordinate];
    }

    xlnt::cell cell(int row, int column)
    {
        return cell(get_column_letter(column + 1) + std::to_string(row + 1));
    }

    int get_highest_row() const
    {
        throw not_implemented();
    }

    int get_highest_column() const
    {
        throw not_implemented();
    }

    std::string calculate_dimension() const
    {
        throw not_implemented();
    }

    range range(const std::string &range_string, int /*row_offset*/, int /*column_offset*/)
    {
        auto colon_index = range_string.find(':');
        if(colon_index != std::string::npos)
        {
            auto min_range = range_string.substr(0, colon_index);
            auto max_range = range_string.substr(colon_index + 1);
            xlnt::range r;
            r.push_back(std::vector<xlnt::cell>());
            r[0].push_back(cell(min_range));
            r[0].push_back(cell(max_range));
            return r;
        }
        throw not_implemented();
    }

    relationship create_relationship(relationship::type relationship_type)
    {
        relationships_.push_back(relationship(relationship_type));
        return relationships_.back();
    }

    //void add_chart(chart chart);

    void merge_cells(const std::string &/*range_string*/)
    {
        throw not_implemented();
    }

    void merge_cells(int /*start_row*/, int /*start_column*/, int /*end_row*/, int /*end_column*/)
    {
        throw not_implemented();
    }

    void unmerge_cells(const std::string &/*range_string*/)
    {
        throw not_implemented();
    }

    void unmerge_cells(int /*start_row*/, int /*start_column*/, int /*end_row*/, int /*end_column*/)
    {
        throw not_implemented();
    }

    void append(const std::vector<xlnt::cell> &cells)
    {
        for(auto cell : cells)
        {
            cells_["a"] = cell;
        }
    }

    void append(const std::unordered_map<std::string, xlnt::cell> &cells)
    {
        for(auto cell : cells)
        {
            cells_[cell.first] = cell.second;
        }
    }

    void append(const std::unordered_map<int, xlnt::cell> &cells)
    {
        for(auto cell : cells)
        {
            cells_[get_column_letter(cell.first + 1)] = cell.second;
        }
    }

    xlnt::range rows() const
    {
        throw not_implemented();
    }

    xlnt::range columns() const
    {
        throw not_implemented();
    }

    void operator=(const worksheet_impl &other) = delete;

    workbook &parent_;
    std::string title_;
    xlnt::cell freeze_panes_;
    std::unordered_map<std::string, xlnt::cell> cells_;
    std::vector<relationship> relationships_;
};

worksheet::worksheet(workbook &parent_workbook, const std::string &title) 
    : impl_(std::make_shared<worksheet_impl>(parent_workbook, title))
{

}

std::string worksheet::to_string() const
{
    return "<Worksheet \"" + impl_->title_ + "\">";
}

workbook worksheet::get_parent() const
{
    return impl_->parent_;
}

void worksheet::garbage_collect()
{
    impl_->garbage_collect();
}

std::set<cell> worksheet::get_cell_collection()
{
    return impl_->get_cell_collection();
}

std::string worksheet::get_title() const
{
    return impl_->title_;
}

void worksheet::set_title(const std::string &title)
{
    impl_->title_ = title;
}

cell worksheet::get_freeze_panes() const
{
    return impl_->freeze_panes_;
}

void worksheet::set_freeze_panes(xlnt::cell top_left_cell)
{
    impl_->set_freeze_panes(top_left_cell);
}

void worksheet::set_freeze_panes(const std::string &top_left_coordinate)
{
    impl_->set_freeze_panes(top_left_coordinate);
}

void worksheet::unfreeze_panes()
{
    impl_->unfreeze_panes();
}

xlnt::cell worksheet::cell(const std::string &coordinate)
{
    return impl_->cell(coordinate);
}

xlnt::cell worksheet::cell(int row, int column)
{
    return impl_->cell(row, column);
}

int worksheet::get_highest_row() const
{
    return impl_->get_highest_row();
}

int worksheet::get_highest_column() const
{
    return impl_->get_highest_column();
}

std::string worksheet::calculate_dimension() const
{
    return impl_->calculate_dimension();
}

range worksheet::range(const std::string &range_string, int row_offset, int column_offset)
{
    return impl_->range(range_string, row_offset, column_offset);
}

relationship worksheet::create_relationship(relationship::type relationship_type)
{
    return impl_->create_relationship(relationship_type);
}

//void worksheet::add_chart(chart chart);

void worksheet::merge_cells(const std::string &range_string)
{
    impl_->merge_cells(range_string);
}

void worksheet::merge_cells(int start_row, int start_column, int end_row, int end_column)
{
    impl_->merge_cells(start_row, start_column, end_row, end_column);
}

void worksheet::unmerge_cells(const std::string &range_string)
{
    impl_->unmerge_cells(range_string);
}

void worksheet::unmerge_cells(int start_row, int start_column, int end_row, int end_column)
{
    impl_->unmerge_cells(start_row, start_column, end_row, end_column);
}

void worksheet::append(const std::vector<xlnt::cell> &cells)
{
    impl_->append(cells);
}

void worksheet::append(const std::unordered_map<std::string, xlnt::cell> &cells)
{
    impl_->append(cells);
}

void worksheet::append(const std::unordered_map<int, xlnt::cell> &cells)
{
    impl_->append(cells);
}

xlnt::range worksheet::rows() const
{
    return impl_->rows();
}

xlnt::range worksheet::columns() const
{
    return impl_->columns();
}

bool worksheet::operator==(const worksheet &other)
{
    return impl_ == other.impl_;
}

void worksheet::operator=(const worksheet &other)
{
    impl_ = other.impl_;
}

cell worksheet::operator[](const std::string &address)
{
    return cell(address);
}

} // namespace xlnt
