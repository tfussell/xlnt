#include <algorithm>

#include "worksheet/worksheet.hpp"
#include "cell/cell.hpp"
#include "cell/value.hpp"
#include "common/datetime.hpp"
#include "worksheet/range.hpp"
#include "worksheet/range_reference.hpp"
#include "common/relationship.hpp"
#include "workbook/workbook.hpp"
#include "detail/worksheet_impl.hpp"
#include "common/exceptions.hpp"

namespace xlnt {

worksheet::worksheet() : d_(nullptr)
{
    
}
    
worksheet::worksheet(detail::worksheet_impl *d) : d_(d)
{
    
}

worksheet::worksheet(const worksheet &rhs) : d_(rhs.d_)
{
    
}
    
worksheet::worksheet(workbook &parent, const std::string &title) : d_(title == "" ? parent.create_sheet().d_ : parent.create_sheet(title).d_)
{
}

bool worksheet::has_frozen_panes() const
{
    return get_frozen_panes() != cell_reference("A1");
}

void worksheet::create_named_range(const std::string &name, const range_reference &reference)
{
    d_->named_ranges_[name] = reference;
}

range worksheet::operator()(const xlnt::cell_reference &top_left, const xlnt::cell_reference &bottom_right)
{
    return get_range({top_left, bottom_right});
}

cell worksheet::operator[](const cell_reference &ref)
{
    return get_cell(ref);
}

std::vector<range_reference> worksheet::get_merged_ranges() const
{
    return d_->merged_cells_;
}

margins &worksheet::get_page_margins()
{
    return d_->page_margins_;
}

void worksheet::auto_filter(const range_reference &reference)
{
    d_->auto_filter_ = reference;
}

void worksheet::auto_filter(const xlnt::range &range)
{
    auto_filter(range.get_reference());
}

range_reference worksheet::get_auto_filter() const
{
    return d_->auto_filter_;
}

bool worksheet::has_auto_filter() const
{
    return d_->auto_filter_.get_width() > 0;
}

void worksheet::unset_auto_filter()
{
    d_->auto_filter_ = range_reference(0, 0, 0, 0);
}

page_setup &worksheet::get_page_setup()
{
    return d_->page_setup_;
}

std::string worksheet::to_string() const
{
    return "<Worksheet \"" + d_->title_ + "\">";
}

workbook &worksheet::get_parent() const
{
    return *d_->parent_;
}

void worksheet::garbage_collect()
{
    auto cell_map_iter = d_->cell_map_.begin();

    while(cell_map_iter != d_->cell_map_.end())
    {
        auto cell_iter = cell_map_iter->second.begin();

        while(cell_iter != cell_map_iter->second.end())
        {
            cell current_cell(&cell_iter->second);

            if(current_cell.garbage_collectible())
            {
                cell_iter = cell_map_iter->second.erase(cell_iter);
                continue;
            }

            cell_iter++;
        }

        if(cell_map_iter->second.empty())
        {
            cell_map_iter = d_->cell_map_.erase(cell_map_iter);
            continue;
        }

        cell_map_iter++;
    }
}

std::list<cell> worksheet::get_cell_collection()
{
    std::list<cell> cells;
    for(auto &c : d_->cell_map_)
    {
        for(auto &d : c.second)
        {
            cells.push_back(cell(&d.second));
        }
    }
    return cells;
}

std::string worksheet::get_title() const
{
    if(d_ == nullptr)
    {
        throw std::runtime_error("null worksheet");
    }
    return d_->title_;
}

void worksheet::set_title(const std::string &title)
{
    d_->title_ = title;
}

cell_reference worksheet::get_frozen_panes() const
{
    return d_->freeze_panes_;
}

void worksheet::freeze_panes(xlnt::cell top_left_cell)
{
    d_->freeze_panes_ = top_left_cell.get_reference();
}

void worksheet::freeze_panes(const std::string &top_left_coordinate)
{
    d_->freeze_panes_ = cell_reference(top_left_coordinate);
}

void worksheet::unfreeze_panes()
{
    d_->freeze_panes_ = cell_reference("A1");
}

const std::unordered_map<column_t, double> &worksheet::get_column_dimensions() const
{
    return d_->column_dimensions_;
}

const std::unordered_map<row_t, double> &worksheet::get_row_dimensions() const
{
    return d_->row_dimensions_;
}

cell worksheet::get_cell(const cell_reference &reference)
{
    if(d_->cell_map_.find(reference.get_row_index()) == d_->cell_map_.end())
    {
        d_->cell_map_[reference.get_row_index()] = std::unordered_map<column_t, detail::cell_impl>();
    }
    
    auto &row = d_->cell_map_[reference.get_row_index()];
    
    if(row.find(reference.get_column_index()) == row.end())
    {
        row[reference.get_column_index()] = detail::cell_impl(d_, reference.get_column_index(), reference.get_row_index());
    }
    
    return cell(&row[reference.get_column_index()]);
}

const cell worksheet::get_cell(const cell_reference &reference) const
{
    return cell(&d_->cell_map_.at(reference.get_row_index()).at(reference.get_column_index()));
}

row_properties &worksheet::get_row_properties(row_t row)
{
    return d_->row_properties_[row];
}

bool worksheet::has_row_properties(row_t row) const
{
    return d_->row_properties_.find(row) != d_->row_properties_.end();
}

range worksheet::get_named_range(const std::string &name)
{
    if(!has_named_range(name))
    {
        throw named_range_exception();
    }
    
    return get_range(d_->named_ranges_[name]);
}

column_t worksheet::get_lowest_column() const
{
    if(d_->cell_map_.empty())
    {
        return 1;
    }
    
    column_t lowest = std::numeric_limits<column_t>::max();
    
    for(auto &row : d_->cell_map_)
    {
        for(auto &c : row.second)
        {
            lowest = std::min(lowest, (column_t)c.first);
        }
    }
    
    return lowest + 1;
}

row_t worksheet::get_lowest_row() const
{
    if(d_->cell_map_.empty())
    {
        return 1;
    }
    
    row_t lowest = std::numeric_limits<row_t>::max();
    
    for(auto &row : d_->cell_map_)
    {
        lowest = std::min(lowest, (row_t)row.first);
    }
    
    return lowest + 1;
}

row_t worksheet::get_highest_row() const
{
    row_t highest = 0;
    
    for(auto &row : d_->cell_map_)
    {
        highest = std::max(highest, (row_t)row.first);
    }
    
    return highest + 1;
}

column_t worksheet::get_highest_column() const
{
    column_t highest = 0;
    
    for(auto &row : d_->cell_map_)
    {
        for(auto &c : row.second)
        {
            highest = std::max(highest, (column_t)c.first);
        }
    }
    
    return highest + 1;
}

range_reference worksheet::calculate_dimension() const
{
    int lowest_column = get_lowest_column();
    int lowest_row = get_lowest_row();
    
    int highest_column = get_highest_column();
    int highest_row = get_highest_row();
    
    return range_reference(lowest_column - 1, lowest_row - 1, highest_column - 1, highest_row - 1);
}

range worksheet::get_range(const range_reference &reference)
{
    return range(*this, reference);
}

const range worksheet::get_range(const range_reference &reference) const
{
    return range(*this, reference);
}

std::vector<relationship> worksheet::get_relationships()
{
    return d_->relationships_;
}

relationship worksheet::create_relationship(relationship::type type, const std::string &target_uri)
{
    std::string r_id = "rId" + std::to_string(d_->relationships_.size() + 1);
    d_->relationships_.push_back(relationship(type, r_id, target_uri));
    return d_->relationships_.back();
}

void worksheet::merge_cells(const range_reference &reference)
{
    d_->merged_cells_.push_back(reference);
    bool first = true;

    for(auto row : get_range(reference))
    {
        for(auto cell : row)
        {
            cell.set_merged(true);

            if(!first)
            {
                if(cell.get_value().is(value::type::string))
                {
                    cell.set_value(value(""));
                }
                else
                {
                    cell.set_value(value::null());
                }
            }

            first = false;
        }
    }
}

void worksheet::unmerge_cells(const range_reference &reference)
{
    auto match = std::find(d_->merged_cells_.begin(), d_->merged_cells_.end(), reference);
    
    if(match == d_->merged_cells_.end())
    {
        throw std::runtime_error("cells not merged");
    }
    
    d_->merged_cells_.erase(match);
    
    for(auto row : get_range(reference))
    {
        for(auto cell : row)
        {
            cell.set_merged(false);
        }
    }
}

void worksheet::append(const std::vector<std::string> &cells)
{
    int row = get_highest_row();
    
    if(d_->cell_map_.size() == 0)
    {
        row--;
    }
    
    int column = 0;
    
    for(auto cell : cells)
    {
        get_cell(cell_reference(column++, row)).set_value(cell);
    }
}

void worksheet::append(const std::vector<int> &cells)
{
    int row = get_highest_row();
    
    if(d_->cell_map_.size() == 0)
    {
        row--;
    }
    
    int column = 0;
    
    for(auto cell : cells)
    {
        get_cell(cell_reference(column++, row)).set_value(value(cell));
    }
}

void worksheet::append(const std::vector<date> &cells)
{
    int row = get_highest_row();
    
    if(d_->cell_map_.size() == 0)
    {
        row--;
    }
    
    int column = 0;
    
    for(auto cell : cells)
    {
        get_cell(cell_reference(column++, row)).set_value(cell);
    }
}

void worksheet::append(const std::unordered_map<std::string, std::string> &cells)
{
	int row = get_highest_row() - 1;

    if(d_->cell_map_.size() != 0)
    {
        row++;
    }

    for(auto cell : cells)
    {
        get_cell(cell_reference(cell.first, row + 1)).set_value(cell.second);
    }
}

void worksheet::append(const std::unordered_map<int, std::string> &cells)
{
    int row = get_highest_row() - 1;

    if(d_->cell_map_.size() != 0)
    {
        row++;
    }

    for(auto cell : cells)
    {
        get_cell(cell_reference(cell.first, row)).set_value(cell.second);
    }
}

xlnt::range worksheet::rows() const
{
    return get_range(calculate_dimension());
}

xlnt::range worksheet::columns() const
{
    return range(*this, calculate_dimension(), major_order::column);
}

bool worksheet::operator==(const worksheet &other) const
{
    return d_ == other.d_;
}

bool worksheet::operator!=(const worksheet &other) const
{
    return d_ != other.d_;
}

bool worksheet::operator==(std::nullptr_t) const
{
    return d_ == nullptr;
}

bool worksheet::operator!=(std::nullptr_t) const
{
    return d_ != nullptr;
}


void worksheet::operator=(const worksheet &other)
{
    d_ = other.d_;
}

const cell worksheet::operator[](const cell_reference &ref) const
{
    return get_cell(ref);
}

range worksheet::operator[](const range_reference &ref)
{
    return get_range(ref);
}

range worksheet::operator[](const std::string &name)
{
    if(has_named_range(name))
    {
        return get_named_range(name);
    }
    return get_range(range_reference(name));
}

bool worksheet::has_named_range(const std::string &name)
{
    return d_->named_ranges_.find(name) != d_->named_ranges_.end();
}

void worksheet::remove_named_range(const std::string &name)
{
    if(!has_named_range(name))
    {
        throw std::runtime_error("worksheet doesn't have named range");
    }

    d_->named_ranges_.erase(name);
}

void worksheet::reserve(std::size_t n)
{
    d_->cell_map_.reserve(n);
}
    
void worksheet::increment_comments()
{
    d_->comment_count_++;
}

void worksheet::decrement_comments()
{
    d_->comment_count_--;
}

std::size_t worksheet::get_comment_count() const
{
    return d_->comment_count_;
}

header_footer &worksheet::get_header_footer()
{
    return d_->header_footer_;
}

const header_footer &worksheet::get_header_footer() const
{
    return d_->header_footer_;
}

header_footer::header_footer()
{
    
}

header::header() : default_(true), font_size_(12)
{
    
}

footer::footer() : default_(true), font_size_(12)
{
    
}

void worksheet::set_parent(xlnt::workbook &wb)
{
    d_->parent_ = &wb;
}

std::vector<std::string> worksheet::get_formula_attributes() const
{
    return {};
}

cell_reference worksheet::get_point_pos(int left, int top) const
{
    static const double DefaultColumnWidth = 51.85;
    static const double DefaultRowHeight = 15.0;

    auto points_to_pixels = [](double value, double dpi) { return (int)std::ceil(value * dpi / 72); };
    
    auto default_height = points_to_pixels(DefaultRowHeight, 96.0);
    auto default_width = points_to_pixels(DefaultColumnWidth, 96.0);

    column_t current_column = 0;
    row_t current_row = 0;

    int left_pos = 0;
    int top_pos = 0;

    while(left_pos <= left)
    {
        current_column++;

        if(d_->column_dimensions_.find(current_column) != d_->column_dimensions_.end())
        {
            auto cdw = d_->column_dimensions_.at(current_column);

            if(cdw >= 0)
            {
                left_pos += points_to_pixels(cdw, 96.0);
                continue;
            }
        }

        left_pos += default_width;
    }

    while(top_pos <= top)
    {
        current_row++;

        if(d_->row_dimensions_.find(current_row) != d_->row_dimensions_.end())
        {
            auto cdh = d_->column_dimensions_.at(current_row);

            if(cdh >= 0)
            {
                top_pos += points_to_pixels(cdh, 96.0);
                continue;
            }
        }

        top_pos += default_height;
    }

    return {current_column - 1, current_row - 1};
}

cell_reference worksheet::get_point_pos(const std::pair<int, int> &point) const
{
    return get_point_pos(point.first, point.second);
}

} // namespace xlnt
