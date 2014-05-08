#include <unordered_map>

#include "worksheet.h"
#include "workbook.h"
#include "cell.h"
#include "packaging/relationship.h"

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

struct worksheet_struct
{
    worksheet_struct(workbook &parent_workbook, worksheet &owner, const std::string &title) 
        : parent_(parent_workbook), title_(title), freeze_panes_(owner, "A", 1), owner_(owner)
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
        if(cell_map_.find(coordinate) == cell_map_.end())
        {
            auto coord = xlnt::cell::coordinate_from_string(coordinate);
            cells_.emplace_back(owner_, coord.column, coord.row);
            cell_map_[coordinate] = &cells_.back();
        }

        return *cell_map_[coordinate];
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

    relationship create_relationship(const std::string &relationship_type)
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
            *cell_map_["a"] = cell;
        }
    }

    void append(const std::unordered_map<std::string, xlnt::cell> &cells)
    {
        for(auto cell : cells)
        {
            *cell_map_[cell.first] = cell.second;
        }
    }

    void append(const std::unordered_map<int, xlnt::cell> &cells)
    {
        for(auto cell : cells)
        {
            *cell_map_[get_column_letter(cell.first + 1)] = cell.second;
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

    void operator=(const worksheet_struct &other) = delete;

    workbook &parent_;
    std::string title_;
    xlnt::cell freeze_panes_;
    std::vector<xlnt::cell> cells_;
    std::unordered_map<std::string, xlnt::cell *> cell_map_;
    std::vector<relationship> relationships_;
    worksheet &owner_;
};

worksheet::worksheet(workbook &parent_workbook, const std::string &title) 
    : root_(nullptr)
{
    root_ = new worksheet_struct(parent_workbook, *this, title);
}

std::string worksheet::to_string() const
{
    return "<Worksheet \"" + root_->title_ + "\">";
}

workbook &worksheet::get_parent() const
{
    return root_->parent_;
}

void worksheet::garbage_collect()
{
    root_->garbage_collect();
}

std::set<cell> worksheet::get_cell_collection()
{
    return root_->get_cell_collection();
}

std::string worksheet::get_title() const
{
    return root_->title_;
}

void worksheet::set_title(const std::string &title)
{
    root_->title_ = title;
}

cell worksheet::get_freeze_panes() const
{
    return root_->freeze_panes_;
}

void worksheet::set_freeze_panes(xlnt::cell top_left_cell)
{
    root_->set_freeze_panes(top_left_cell);
}

void worksheet::set_freeze_panes(const std::string &top_left_coordinate)
{
    root_->set_freeze_panes(top_left_coordinate);
}

void worksheet::unfreeze_panes()
{
    root_->unfreeze_panes();
}

xlnt::cell worksheet::cell(const std::string &coordinate)
{
    return root_->cell(coordinate);
}

xlnt::cell worksheet::cell(int row, int column)
{
    return root_->cell(row, column);
}

int worksheet::get_highest_row() const
{
    return root_->get_highest_row();
}

int worksheet::get_highest_column() const
{
    return root_->get_highest_column();
}

std::string worksheet::calculate_dimension() const
{
    return root_->calculate_dimension();
}

range worksheet::range(const std::string &range_string, int row_offset, int column_offset)
{
    return root_->range(range_string, row_offset, column_offset);
}

relationship worksheet::create_relationship(const std::string &relationship_type)
{
    return root_->create_relationship(relationship_type);
}

//void worksheet::add_chart(chart chart);

void worksheet::merge_cells(const std::string &range_string)
{
    root_->merge_cells(range_string);
}

void worksheet::merge_cells(int start_row, int start_column, int end_row, int end_column)
{
    root_->merge_cells(start_row, start_column, end_row, end_column);
}

void worksheet::unmerge_cells(const std::string &range_string)
{
    root_->unmerge_cells(range_string);
}

void worksheet::unmerge_cells(int start_row, int start_column, int end_row, int end_column)
{
    root_->unmerge_cells(start_row, start_column, end_row, end_column);
}

void worksheet::append(const std::vector<xlnt::cell> &cells)
{
    root_->append(cells);
}

void worksheet::append(const std::unordered_map<std::string, xlnt::cell> &cells)
{
    root_->append(cells);
}

void worksheet::append(const std::unordered_map<int, xlnt::cell> &cells)
{
    root_->append(cells);
}

xlnt::range worksheet::rows() const
{
    return root_->rows();
}

xlnt::range worksheet::columns() const
{
    return root_->columns();
}

bool worksheet::operator==(const worksheet &other)
{
    return root_ == other.root_;
}

void worksheet::operator=(const worksheet &other)
{
    root_ = other.root_;
}

cell worksheet::operator[](const std::string &address)
{
    return cell(address);
}

} // namespace xlnt
