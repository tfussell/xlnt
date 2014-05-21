#include "worksheet.h"
#include "cell.h"
#include "named_range.h"
#include "reference.h"
#include "relationship.h"
#include "workbook.h"

namespace xlnt {
    
struct worksheet_struct
{
    worksheet_struct(workbook &parent_workbook, const std::string &title)
    : parent_(parent_workbook), title_(title), freeze_panes_("A1")
    {
        
    }
    
    ~worksheet_struct()
    {
        clear();
    }
    
    void clear()
    {
        while(!cell_map_.empty())
        {
            cell::deallocate(cell_map_.begin()->second.root_);
            cell_map_.erase(cell_map_.begin()->first);
        }
    }
    
    void garbage_collect()
    {
        std::vector<cell_reference> null_cell_keys;
        
        for(auto key_cell_pair : cell_map_)
        {
            if(key_cell_pair.second.get_data_type() == cell::type::null)
            {
                null_cell_keys.push_back(key_cell_pair.first);
            }
        }
        
        while(!null_cell_keys.empty())
        {
            cell_map_.erase(null_cell_keys.back());
            null_cell_keys.pop_back();
        }
    }
    
    std::list<cell> get_cell_collection()
    {
        std::list<xlnt::cell> cells;
        for(auto cell : cell_map_)
        {
            cells.push_front(xlnt::cell(cell.second));
        }
        return cells;
    }
    
    std::string get_title() const
    {
        return title_;
    }
    
    void set_title(const std::string &title)
    {
        title_ = title;
    }
    
    cell_reference get_freeze_panes() const
    {
        return freeze_panes_;
    }
    
    void set_freeze_panes(const cell_reference &ref)
    {
        freeze_panes_ = ref;
    }
    
    void unfreeze_panes()
    {
        freeze_panes_ = "A1";
    }
    
    xlnt::cell get_cell(const cell_reference &reference)
    {
        if(cell_map_.find(reference) == cell_map_.end())
        {
            auto cell = cell::allocate(this, reference.get_column_index(), reference.get_row_index());
            cell_map_[reference] = cell;
        }
        
        return cell_map_[reference];
    }
    
    int get_highest_row() const
    {
        int highest = 0;
        for(auto cell : cell_map_)
        {
            highest = (std::max)(highest, cell.first.get_row_index());
        }
        return highest + 1;
    }
    
    int get_highest_column() const
    {
        int highest = 0;
        for(auto cell : cell_map_)
        {
            highest = (std::max)(highest, cell.first.get_column_index());
        }
        return highest + 1;
    }
    
    int get_lowest_row() const
    {
        if(cell_map_.empty())
        {
            return 1;
        }
        
        int lowest = cell_map_.begin()->first.get_row_index();
        
        for(auto cell : cell_map_)
        {
            lowest = (std::min)(lowest, cell.first.get_row_index());
        }
        
        return lowest + 1;
    }
    
    int get_lowest_column() const
    {
        if(cell_map_.empty())
        {
            return 1;
        }
        
        int lowest = cell_map_.begin()->first.get_column_index();
        
        for(auto cell : cell_map_)
        {
            lowest = (std::min)(lowest, cell.first.get_column_index());
        }
        
        return lowest + 1;
    }
    
    range_reference calculate_dimension() const
    {
        int lowest_column = get_lowest_column();
        int lowest_row = get_lowest_row();
        
        int highest_column = get_highest_column();
        int highest_row = get_highest_row();
        
        return range_reference(lowest_column - 1, lowest_row - 1, highest_column - 1, highest_row - 1);
    }
    
    xlnt::range get_range(const range_reference &reference)
    {
        xlnt::range r;
        
        if(!reference.is_single_cell())
        {
            cell_reference min_ref = reference.get_top_left();
            cell_reference max_ref = reference.get_bottom_right();
            
            for(int row = min_ref.get_row_index(); row <= max_ref.get_row_index(); row++)
            {
                r.push_back(std::vector<xlnt::cell>());
                for(int column = min_ref.get_column_index(); column <= max_ref.get_column_index(); column++)
                {
                    r.back().push_back(get_cell(cell_reference(column, row)));
                }
            }
        }
        else
        {
            r.push_back(std::vector<xlnt::cell>());
            r.back().push_back(get_cell(reference.get_top_left()));
        }
        
        return r;
    }
    
    relationship create_relationship(const std::string &relationship_type, const std::string &target_uri)
    {
        std::string r_id = "rId" + std::to_string(relationships_.size() + 1);
        relationships_.push_back(relationship(relationship_type, r_id, target_uri));
        return relationships_.back();
    }
    
    //void add_chart(chart chart);
    
    void merge_cells(const range_reference &reference)
    {
        merged_cells_.push_back(reference);
        bool first = true;
        for(auto row : get_range(reference))
        {
            for(auto cell : row)
            {
                cell.set_merged(true);
                if(!first)
                {
                    cell.bind_value();
                }
                first = false;
            }
        }
    }
    
    void unmerge_cells(const range_reference &reference)
    {
        auto match = std::find(merged_cells_.begin(), merged_cells_.end(), reference);
        if(match == merged_cells_.end())
        {
            throw std::runtime_error("cells not merged");
        }
        merged_cells_.erase(match);
        for(auto row : get_range(reference))
        {
            for(auto cell : row)
            {
                cell.set_merged(false);
            }
        }
    }
    
    void append(const std::vector<std::string> &cells)
    {
        int row = get_highest_row();
        if(cell_map_.size() == 0)
        {
            row--;
        }
        int column = 0;
        for(auto cell : cells)
        {
            this->get_cell(cell_reference(column++, row)) = cell;
        }
    }
    
    void append(const std::unordered_map<std::string, std::string> &cells)
    {
        int row = get_highest_row() - 1;
        if(cell_map_.size() != 0)
        {
            row++;
        }
        for(auto cell : cells)
        {
            this->get_cell(cell_reference(cell.first, row + 1)) = cell.second;
        }
    }
    
    void append(const std::unordered_map<int, std::string> &cells)
    {
        int row = get_highest_row() - 1;
        if(cell_map_.size() != 0)
        {
            row++;
        }
        for(auto cell : cells)
        {
            this->get_cell(cell_reference(cell.first, row)) = cell.second;
        }
    }
    
    xlnt::range rows()
    {
        return get_range(calculate_dimension());
    }
    
    xlnt::range columns()
    {
        int max_row = get_highest_row();
        xlnt::range cols;
        for(int col_idx = 0; col_idx < get_highest_column(); col_idx++)
        {
            cols.push_back(std::vector<xlnt::cell>());
            for(auto row : get_range(range_reference(col_idx, 0, col_idx, max_row)))
            {
                cols.back().push_back(row[0]);
            }
        }
        return cols;
    }
    
    void operator=(const worksheet_struct &other) = delete;
    
    workbook &parent_;
    std::string title_;
    cell_reference freeze_panes_;
    std::unordered_map<cell_reference, xlnt::cell, cell_reference_hash> cell_map_;
    std::vector<relationship> relationships_;
    page_setup page_setup_;
    range_reference auto_filter_;
    margins page_margins_;
    std::vector<range_reference> merged_cells_;
    std::unordered_map<std::string, named_range> named_ranges_;
};

worksheet::worksheet() : root_(nullptr)
{
    
}

worksheet::worksheet(worksheet_struct *root) : root_(root)
{
}

worksheet::worksheet(workbook &parent)
{
    *this = parent.create_sheet();
}

bool worksheet::has_frozen_panes() const
{
    return !(get_frozen_panes() == cell_reference("A1"));
}

named_range worksheet::create_named_range(const std::string &name, const range_reference &reference)
{
    root_->named_ranges_[name] = named_range(named_range::allocate(name, *this, reference));
    return root_->named_ranges_[name];
}

std::vector<range_reference> worksheet::get_merged_ranges() const
{
    return root_->merged_cells_;
}

margins &worksheet::get_page_margins()
{
    return root_->page_margins_;
}

void worksheet::set_auto_filter(const range_reference &reference)
{
    root_->auto_filter_ = reference;
}

void worksheet::set_auto_filter(const xlnt::range &range)
{
    set_auto_filter(range_reference(range[0][0].get_reference(), range.back().back().get_reference()));
}

range_reference worksheet::get_auto_filter() const
{
    return root_->auto_filter_;
}

bool worksheet::has_auto_filter() const
{
    return root_->auto_filter_.get_width() > 0;
}

void worksheet::unset_auto_filter()
{
    root_->auto_filter_ = range_reference(0, 0, 0, 0);
}

page_setup &worksheet::get_page_setup()
{
    return root_->page_setup_;
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

std::list<cell> worksheet::get_cell_collection()
{
    return root_->get_cell_collection();
}

std::string worksheet::get_title() const
{
    if(root_ == nullptr)
    {
        throw std::runtime_error("null worksheet");
    }
    return root_->title_;
}

void worksheet::set_title(const std::string &title)
{
    root_->title_ = title;
}

cell_reference worksheet::get_frozen_panes() const
{
    return root_->freeze_panes_;
}

void worksheet::freeze_panes(xlnt::cell top_left_cell)
{
    root_->set_freeze_panes(top_left_cell.get_reference());
}

void worksheet::freeze_panes(const std::string &top_left_coordinate)
{
    root_->set_freeze_panes(top_left_coordinate);
}

void worksheet::unfreeze_panes()
{
    root_->unfreeze_panes();
}

cell worksheet::get_cell(const cell_reference &reference)
{
    return root_->get_cell(reference);
}

range worksheet::get_named_range(const std::string &name)
{
    return get_range(root_->parent_.get_named_range(name, *this).get_target_range());
}

int worksheet::get_highest_row() const
{
    return root_->get_highest_row();
}

int worksheet::get_highest_column() const
{
    return root_->get_highest_column();
}

range_reference worksheet::calculate_dimension() const
{
    return root_->calculate_dimension();
}

range worksheet::get_range(const range_reference &reference)
{
    return root_->get_range(reference);
}

std::vector<relationship> worksheet::get_relationships()
{
    return root_->relationships_;
}

relationship worksheet::create_relationship(const std::string &relationship_type, const std::string &target_uri)
{
    return root_->create_relationship(relationship_type, target_uri);
}

//void worksheet::add_chart(chart chart);

void worksheet::merge_cells(const range_reference &reference)
{
    root_->merge_cells(reference);
}

void worksheet::unmerge_cells(const range_reference &reference)
{
    root_->unmerge_cells(reference);
}

void worksheet::append(const std::vector<std::string> &cells)
{
    root_->append(cells);
}

void worksheet::append(const std::unordered_map<std::string, std::string> &cells)
{
    root_->append(cells);
}

void worksheet::append(const std::unordered_map<int, std::string> &cells)
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

bool worksheet::operator==(const worksheet &other) const
{
    return root_ == other.root_;
}

bool worksheet::operator!=(const worksheet &other) const
{
    return root_ != other.root_;
}

bool worksheet::operator==(std::nullptr_t) const
{
    return root_ == nullptr;
}

bool worksheet::operator!=(std::nullptr_t) const
{
    return root_ != nullptr;
}


void worksheet::operator=(const worksheet &other)
{
    root_ = other.root_;
}

cell worksheet::operator[](const cell_reference &ref)
{
    return get_cell(ref);
}

range worksheet::operator[](const range_reference &ref)
{
    return get_range(ref);
}

range worksheet::operator[](const std::string &name)
{
    if(root_->parent_.has_named_range(name, *this))
    {
        return get_named_range(name);
    }
    return get_range(range_reference(name));
}

worksheet worksheet::allocate(xlnt::workbook &owner, const std::string &title)
{
    return new worksheet_struct(owner, title);
}
    
void worksheet::deallocate(xlnt::worksheet ws)
{
    delete ws.root_;
}
    
} // namespace xlnt
