// Copyright (c) 2014-2016 Thomas Fussell
// Copyright (c) 2010-2015 openpyxl
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, WRISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE
//
// @license: http://www.opensource.org/licenses/mit-license.php
// @author: see AUTHORS file
#include <algorithm>
#include <cmath>
#include <limits>

#include <xlnt/cell/cell.hpp>
#include <xlnt/cell/cell_reference.hpp>
#include <xlnt/cell/index_types.hpp>
#include <xlnt/packaging/relationship.hpp>
#include <xlnt/utils/date.hpp>
#include <xlnt/utils/datetime.hpp>
#include <xlnt/utils/exceptions.hpp>
#include <xlnt/workbook/named_range.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/workbook/worksheet_iterator.hpp>
#include <xlnt/worksheet/cell_iterator.hpp>
#include <xlnt/worksheet/const_cell_iterator.hpp>
#include <xlnt/worksheet/const_range_iterator.hpp>
#include <xlnt/worksheet/header_footer.hpp>
#include <xlnt/worksheet/range.hpp>
#include <xlnt/worksheet/range_iterator.hpp>
#include <xlnt/worksheet/range_reference.hpp>
#include <xlnt/worksheet/worksheet.hpp>

#include <detail/cell_impl.hpp>
#include <detail/constants.hpp>
#include <detail/workbook_impl.hpp>
#include <detail/worksheet_impl.hpp>

namespace {

int points_to_pixels(double points, double dpi)
{
    return static_cast<int>(std::ceil(points * dpi / 72));
};

} // namespace

namespace xlnt {

worksheet::worksheet()
    : d_(nullptr)
{
}

worksheet::worksheet(detail::worksheet_impl *d)
    : d_(d)
{
}

worksheet::worksheet(const worksheet &rhs)
    : d_(rhs.d_)
{
}

bool worksheet::has_frozen_panes() const
{
    return !d_->views_.empty() && d_->views_.front().has_pane()
        && (d_->views_.front().pane().state == pane_state::frozen
               || d_->views_.front().pane().state == pane_state::frozen_split);
}

std::string worksheet::unique_sheet_name(const std::string &value) const
{
    auto names = workbook().sheet_titles();
    auto match = std::find(names.begin(), names.end(), value);

    std::size_t append = 0;

    while (match != names.end())
    {
        append++;
        match = std::find(names.begin(), names.end(), value + std::to_string(append));
    }

    return append == 0 ? value : value + std::to_string(append);
}

void worksheet::create_named_range(const std::string &name, const std::string &reference_string)
{
    create_named_range(name, range_reference(reference_string));
}

void worksheet::create_named_range(const std::string &name, const range_reference &reference)
{
    try
    {
        auto temp = cell_reference::split_reference(name);

        // name is a valid reference, make sure it's outside the allowed range

        if (column_t(temp.first).index <= column_t("XFD").index && temp.second <= 1048576)
        {
            throw invalid_parameter(); //("named range name must be outside the range A1-XFD1048576");
        }
    }
    catch (xlnt::invalid_cell_reference)
    {
        // name is not a valid reference, that's good
    }

    std::vector<named_range::target> targets;
    targets.push_back({*this, reference});

    d_->named_ranges_[name] = xlnt::named_range(name, targets);
}

range worksheet::operator()(const xlnt::cell_reference &top_left, const xlnt::cell_reference &bottom_right)
{
    return range(range_reference(top_left, bottom_right));
}

cell worksheet::operator[](const cell_reference &ref)
{
    return cell(ref);
}

std::vector<range_reference> worksheet::merged_ranges() const
{
    return d_->merged_cells_;
}

bool worksheet::has_page_margins() const
{
    return d_->page_margins_.is_set();
}

bool worksheet::has_page_setup() const
{
    return d_->page_setup_.is_set();
}

page_margins worksheet::page_margins() const
{
    return d_->page_margins_.get();
}

void worksheet::page_margins(const class page_margins &margins)
{
    d_->page_margins_ = margins;
}

void worksheet::auto_filter(const std::string &reference_string)
{
    auto_filter(range_reference(reference_string));
}

void worksheet::auto_filter(const range_reference &reference)
{
    d_->auto_filter_ = reference;
}

void worksheet::auto_filter(const xlnt::range &range)
{
    auto_filter(range.reference());
}

range_reference worksheet::auto_filter() const
{
    return d_->auto_filter_.get();
}

bool worksheet::has_auto_filter() const
{
    return d_->auto_filter_.is_set();
}

void worksheet::clear_auto_filter()
{
    d_->auto_filter_.clear();
}

void worksheet::page_setup(const struct page_setup &setup)
{
    d_->page_setup_ = setup;
}

page_setup worksheet::page_setup() const
{
    if (!has_page_setup())
    {
        throw invalid_attribute();
    }

    return d_->page_setup_.get();
}

workbook &worksheet::workbook()
{
    return *d_->parent_;
}

const workbook &worksheet::workbook() const
{
    return *d_->parent_;
}

void worksheet::garbage_collect()
{
    auto cell_map_iter = d_->cell_map_.begin();

    while (cell_map_iter != d_->cell_map_.end())
    {
        auto cell_iter = cell_map_iter->second.begin();

        while (cell_iter != cell_map_iter->second.end())
        {
            class cell current_cell(&cell_iter->second);

            if (current_cell.garbage_collectible())
            {
                cell_iter = cell_map_iter->second.erase(cell_iter);
                continue;
            }

            cell_iter++;
        }

        if (cell_map_iter->second.empty())
        {
            cell_map_iter = d_->cell_map_.erase(cell_map_iter);
            continue;
        }

        cell_map_iter++;
    }
}

void worksheet::id(std::size_t id)
{
    d_->id_ = id;
}

std::size_t worksheet::id() const
{
    return d_->id_;
}

std::string worksheet::title() const
{
    return d_->title_;
}

void worksheet::title(const std::string &title)
{
    if (title.length() > 31)
    {
        throw invalid_sheet_title(title);
    }

    if (title.find_first_of("*:/\\?[]") != std::string::npos)
    {
        throw invalid_sheet_title(title);
    }

    auto same_title =
        std::find_if(workbook().begin(), workbook().end(), [&](worksheet ws) { return ws.title() == title; });

    if (same_title != workbook().end() && *same_title != *this)
    {
        throw invalid_sheet_title(title);
    }

    workbook().d_->sheet_title_rel_id_map_[title] = workbook().d_->sheet_title_rel_id_map_[d_->title_];
    workbook().d_->sheet_title_rel_id_map_.erase(d_->title_);
    d_->title_ = title;
}

cell_reference worksheet::frozen_panes() const
{
    if (!has_frozen_panes())
    {
        throw xlnt::invalid_attribute();
    }

    return d_->views_.front().pane().top_left_cell.get();
}

void worksheet::freeze_panes(xlnt::cell top_left_cell)
{
    freeze_panes(top_left_cell.reference());
}

void worksheet::freeze_panes(const cell_reference &ref)
{
    if (!has_view())
    {
        d_->views_.push_back(sheet_view());
    }

    auto &primary_view = d_->views_.front();

    if (!primary_view.has_pane())
    {
        primary_view.pane(pane());
    }

    primary_view.pane().top_left_cell = ref;
    primary_view.pane().state = pane_state::frozen;

    primary_view.clear_selections();
    primary_view.add_selection(selection());

    if (ref == "A1")
    {
        unfreeze_panes();
    }
    else if (ref.column() == "A")
    {
        primary_view.add_selection(selection());
        primary_view.selection(0).pane(pane_corner::bottom_left);
        primary_view.selection(0).active_cell(ref.make_offset(0, -1)); // cell above
        primary_view.selection(1).active_cell(ref);
        primary_view.pane().active_pane = pane_corner::bottom_left;
        primary_view.pane().y_split = ref.row() - 1;
    }
    else if (ref.row() == 1)
    {
        primary_view.add_selection(selection());
        primary_view.selection(0).pane(pane_corner::top_right);
        primary_view.selection(0).active_cell(ref.make_offset(-1, 0)); // cell to the left
        primary_view.selection(1).active_cell(ref);
        primary_view.pane().active_pane = pane_corner::top_right;
        primary_view.pane().x_split = ref.column_index() - 1;
    }
    else
    {
        primary_view.add_selection(selection());
        primary_view.add_selection(selection());
        primary_view.selection(0).pane(pane_corner::top_right);
        primary_view.selection(0).active_cell(ref.make_offset(0, -1)); // cell above
        primary_view.selection(1).pane(pane_corner::bottom_left);
        primary_view.selection(1).active_cell(ref.make_offset(-1, 0)); // cell to the left
        primary_view.selection(2).pane(pane_corner::bottom_right);
        primary_view.selection(2).active_cell(ref);
        primary_view.pane().active_pane = pane_corner::bottom_right;
        primary_view.pane().x_split = ref.column_index() - 1;
        primary_view.pane().y_split = ref.row() - 1;
    }
}

void worksheet::unfreeze_panes()
{
    if (!has_view()) return;

    auto &primary_view = d_->views_.front();

    primary_view.clear_selections();
    primary_view.clear_pane();
}

cell worksheet::cell(column_t column, row_t row)
{
    return cell(cell_reference(column, row));
}

const cell worksheet::cell(column_t column, row_t row) const
{
    return cell(cell_reference(column, row));
}

cell worksheet::cell(const cell_reference &reference)
{
    if (d_->cell_map_.find(reference.row()) == d_->cell_map_.end())
    {
        d_->cell_map_[reference.row()] = std::unordered_map<column_t, detail::cell_impl>();
    }

    auto &row = d_->cell_map_[reference.row()];

    if (row.find(reference.column_index()) == row.end())
    {
        auto &impl = row[reference.column_index()] = detail::cell_impl();
        impl.parent_ = d_;
        impl.column_ = reference.column_index();
        impl.row_ = reference.row();
    }

    return xlnt::cell(&row[reference.column_index()]);
}

const cell worksheet::cell(const cell_reference &reference) const
{
    return xlnt::cell(&d_->cell_map_.at(reference.row()).at(reference.column_index()));
}

bool worksheet::has_cell(const cell_reference &reference) const
{
    const auto row = d_->cell_map_.find(reference.row());
    if (row == d_->cell_map_.cend()) return false;

    const auto col = row->second.find(reference.column_index());
    if (col == row->second.cend()) return false;

    return true;
}

bool worksheet::has_row_properties(row_t row) const
{
    return d_->row_properties_.find(row) != d_->row_properties_.end();
}

range worksheet::named_range(const std::string &name)
{
    if (!workbook().has_named_range(name))
    {
        throw key_not_found();
    }

    if (!has_named_range(name))
    {
        throw key_not_found();
    }

    return range(d_->named_ranges_[name].targets()[0].second);
}

column_t worksheet::lowest_column() const
{
    if (d_->cell_map_.empty())
    {
        return constants::min_column();
    }

    column_t lowest = constants::max_column();

    for (auto &row : d_->cell_map_)
    {
        for (auto &c : row.second)
        {
            lowest = std::min(lowest, c.first);
        }
    }

    return lowest;
}

row_t worksheet::lowest_row() const
{
    if (d_->cell_map_.empty())
    {
        return constants::min_row();
    }

    row_t lowest = constants::max_row();

    for (auto &row : d_->cell_map_)
    {
        lowest = std::min(lowest, row.first);
    }

    return lowest;
}

row_t worksheet::highest_row() const
{
    row_t highest = constants::min_row();

    for (auto &row : d_->cell_map_)
    {
        highest = std::max(highest, row.first);
    }

    return highest;
}

column_t worksheet::highest_column() const
{
    column_t highest = constants::min_column();

    for (auto &row : d_->cell_map_)
    {
        for (auto &c : row.second)
        {
            highest = std::max(highest, c.first);
        }
    }

    return highest;
}

range_reference worksheet::calculate_dimension() const
{
    return range_reference(lowest_column(), lowest_row(), highest_column(), highest_row());
}

range worksheet::range(const std::string &reference_string)
{
    return range(range_reference(reference_string));
}

range worksheet::range(const range_reference &reference)
{
    return xlnt::range(*this, reference);
}

const range worksheet::range(const std::string &reference_string) const
{
    return range(range_reference(reference_string));
}

const range worksheet::range(const range_reference &reference) const
{
    return xlnt::range(*this, reference);
}

void worksheet::merge_cells(const std::string &reference_string)
{
    merge_cells(range_reference(reference_string));
}

void worksheet::unmerge_cells(const std::string &reference_string)
{
    unmerge_cells(range_reference(reference_string));
}

void worksheet::merge_cells(const range_reference &reference)
{
    d_->merged_cells_.push_back(reference);
    bool first = true;

    for (auto row : range(reference))
    {
        for (auto cell : row)
        {
            cell.merged(true);

            if (!first)
            {
                if (cell.data_type() == cell::type::string)
                {
                    cell.value("");
                }
                else
                {
                    cell.clear_value();
                }
            }

            first = false;
        }
    }
}

void worksheet::merge_cells(column_t start_column, row_t start_row, column_t end_column, row_t end_row)
{
    merge_cells(xlnt::range_reference(start_column, start_row, end_column, end_row));
}

void worksheet::unmerge_cells(const range_reference &reference)
{
    auto match = std::find(d_->merged_cells_.begin(), d_->merged_cells_.end(), reference);

    if (match == d_->merged_cells_.end())
    {
        throw invalid_parameter();
    }

    d_->merged_cells_.erase(match);

    for (auto row : range(reference))
    {
        for (auto cell : row)
        {
            cell.merged(false);
        }
    }
}

void worksheet::unmerge_cells(column_t start_column, row_t start_row, column_t end_column, row_t end_row)
{
    unmerge_cells(xlnt::range_reference(start_column, start_row, end_column, end_row));
}

void worksheet::append()
{
    cell(cell_reference(1, next_row()));
}

void worksheet::append(const std::vector<std::string> &cells)
{
    xlnt::cell_reference next(1, next_row());

    for (auto cell : cells)
    {
        this->cell(next).value(cell);
        next.column_index(next.column_index() + 1);
    }
}

row_t worksheet::next_row() const
{
    auto row = highest_row() + 1;

    if (row == 2 && d_->cell_map_.size() == 0)
    {
        row = 1;
    }

    return row;
}

void worksheet::append(const std::vector<int> &cells)
{
    xlnt::cell_reference next(1, next_row());

    for (auto cell : cells)
    {
        this->cell(next).value(cell);
        next.column_index(next.column_index() + 1);
    }
}

void worksheet::append(const std::unordered_map<std::string, std::string> &cells)
{
    auto row = next_row();

    for (auto cell : cells)
    {
        this->cell(cell_reference(cell.first, row)).value(cell.second);
    }
}

void worksheet::append(const std::unordered_map<int, std::string> &cells)
{
    auto row = next_row();

    for (auto cell : cells)
    {
        this->cell(cell_reference(static_cast<column_t::index_t>(cell.first), row)).value(cell.second);
    }
}

void worksheet::append(const std::vector<int>::const_iterator begin, const std::vector<int>::const_iterator end)
{
    xlnt::cell_reference next(1, next_row());

    for (auto i = begin; i != end; i++)
    {
        cell(next).value(*i);
        next.column_index(next.column_index() + 1);
    }
}

xlnt::range worksheet::rows() const
{
    return range(calculate_dimension());
}

xlnt::range worksheet::rows(const std::string &range_string) const
{
    return range(range_reference(range_string));
}

xlnt::range worksheet::rows(const std::string &range_string, int row_offset, int column_offset) const
{
    range_reference reference(range_string);
    return range(reference.make_offset(column_offset, row_offset));
}

xlnt::range worksheet::rows(int row_offset, int column_offset) const
{
    range_reference reference(calculate_dimension());
    return range(reference.make_offset(column_offset, row_offset));
}

xlnt::range worksheet::columns() const
{
    return xlnt::range(*this, calculate_dimension(), major_order::column);
}

bool worksheet::operator==(const worksheet &other) const
{
    return compare(other, true);
}

bool worksheet::compare(const worksheet &other, bool reference) const
{
    if (reference)
    {
        return d_ == other.d_;
    }

    if (d_->parent_ != other.d_->parent_) return false;

    for (auto &row : d_->cell_map_)
    {
        if (other.d_->cell_map_.find(row.first) == other.d_->cell_map_.end())
        {
            return false;
        }

        for (auto &cell : row.second)
        {
            if (other.d_->cell_map_[row.first].find(cell.first) == other.d_->cell_map_[row.first].end())
            {
                return false;
            }

            xlnt::cell this_cell(&cell.second);
            xlnt::cell other_cell(&other.d_->cell_map_[row.first][cell.first]);

            if (this_cell.data_type() != other_cell.data_type())
            {
                return false;
            }

            if (this_cell.data_type() == xlnt::cell::type::numeric
                && std::fabs(this_cell.value<long double>() - other_cell.value<long double>()) > 0.L)
            {
                return false;
            }
        }
    }

    // todo: missing some comparisons

    if (d_->auto_filter_ == other.d_->auto_filter_ && d_->views_ == other.d_->views_
        && d_->merged_cells_ == other.d_->merged_cells_)
    {
        return true;
    }

    return false;
}

bool worksheet::operator!=(const worksheet &other) const
{
    return !(*this == other);
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
    return cell(ref);
}

range worksheet::operator[](const range_reference &ref)
{
    return range(ref);
}

range worksheet::operator[](const std::string &name)
{
    if (has_named_range(name))
    {
        return named_range(name);
    }
    return range(range_reference(name));
}

bool worksheet::has_named_range(const std::string &name)
{
    return d_->named_ranges_.find(name) != d_->named_ranges_.end();
}

void worksheet::remove_named_range(const std::string &name)
{
    if (!has_named_range(name))
    {
        throw key_not_found();
    }

    d_->named_ranges_.erase(name);
}

void worksheet::reserve(std::size_t n)
{
    d_->cell_map_.reserve(n);
}

class header_footer worksheet::header_footer() const
{
    return d_->header_footer_.get();
}

void worksheet::parent(xlnt::workbook &wb)
{
    d_->parent_ = &wb;
}

std::vector<std::string> worksheet::formula_attributes() const
{
    return {};
}

cell_reference worksheet::point_pos(int left, int top) const
{
    column_t current_column = 1;
    row_t current_row = 1;

    int left_pos = 0;
    int top_pos = 0;

    while (left_pos <= left)
    {
        left_pos += column_width(current_column++);
    }

    while (top_pos <= top)
    {
        top_pos += row_height(current_row++);
    }

    return {current_column - 1, current_row - 1};
}

cell_reference worksheet::point_pos(const std::pair<int, int> &point) const
{
    return point_pos(point.first, point.second);
}

void worksheet::sheet_state(xlnt::sheet_state state)
{
    page_setup().sheet_state(state);
}

sheet_state worksheet::sheet_state() const
{
    return page_setup().sheet_state();
}

void worksheet::add_column_properties(column_t column, const xlnt::column_properties &props)
{
    d_->column_properties_[column] = props;
}

bool worksheet::has_column_properties(column_t column) const
{
    return d_->column_properties_.find(column) != d_->column_properties_.end();
}

column_properties &worksheet::column_properties(column_t column)
{
    return d_->column_properties_[column];
}

const column_properties &worksheet::column_properties(column_t column) const
{
    return d_->column_properties_.at(column);
}

row_properties &worksheet::row_properties(row_t row)
{
    return d_->row_properties_[row];
}

const row_properties &worksheet::row_properties(row_t row) const
{
    return d_->row_properties_.at(row);
}

worksheet::iterator worksheet::begin()
{
    auto dimensions = calculate_dimension();
    cell_reference top_right(dimensions.bottom_right().column_index(), dimensions.top_left().row());
    range_reference row_range(dimensions.top_left(), top_right);

    return iterator(*this, row_range, dimensions, major_order::row);
}

worksheet::iterator worksheet::end()
{
    auto dimensions = calculate_dimension();
    auto past_end_row_index = dimensions.bottom_right().row() + 1;
    cell_reference bottom_left(dimensions.top_left().column_index(), past_end_row_index);
    cell_reference bottom_right(dimensions.bottom_right().column_index(), past_end_row_index);

    return iterator(*this, range_reference(bottom_left, bottom_right), dimensions, major_order::row);
}

worksheet::const_iterator worksheet::cbegin() const
{
    auto dimensions = calculate_dimension();
    cell_reference top_right(dimensions.bottom_right().column_index(), dimensions.top_left().row());
    range_reference row_range(dimensions.top_left(), top_right);

    return const_iterator(*this, row_range, major_order::row);
}

worksheet::const_iterator worksheet::cend() const
{
    auto dimensions = calculate_dimension();
    auto past_end_row_index = dimensions.bottom_right().row() + 1;
    cell_reference bottom_left(dimensions.top_left().column_index(), past_end_row_index);
    cell_reference bottom_right(dimensions.bottom_right().column_index(), past_end_row_index);

    return const_iterator(*this, range_reference(bottom_left, bottom_right), major_order::row);
}

worksheet::const_iterator worksheet::begin() const
{
    return cbegin();
}

worksheet::const_iterator worksheet::end() const
{
    return cend();
}

range worksheet::iter_cells(bool skip_null)
{
    return xlnt::range(*this, calculate_dimension(), major_order::row, skip_null);
}

void worksheet::print_title_rows(row_t last_row)
{
    print_title_rows(1, last_row);
}

void worksheet::print_title_rows(row_t first_row, row_t last_row)
{
    d_->print_title_rows_ = std::to_string(first_row) + ":" + std::to_string(last_row);
}

void worksheet::print_title_cols(column_t last_column)
{
    print_title_cols(1, last_column);
}

void worksheet::print_title_cols(column_t first_column, column_t last_column)
{
    d_->print_title_cols_ = first_column.column_string() + ":" + last_column.column_string();
}

std::string worksheet::print_titles() const
{
    if (!d_->print_title_rows_.empty() && !d_->print_title_cols_.empty())
    {
        return d_->title_ + "!" + d_->print_title_rows_ + "," + d_->title_ + "!" + d_->print_title_cols_;
    }
    else if (!d_->print_title_cols_.empty())
    {
        return d_->title_ + "!" + d_->print_title_cols_;
    }
    else
    {
        return d_->title_ + "!" + d_->print_title_rows_;
    }
}

void worksheet::print_area(const std::string &print_area)
{
    d_->print_area_ = range_reference::make_absolute(range_reference(print_area));
}

range_reference worksheet::print_area() const
{
    return d_->print_area_.get();
}

bool worksheet::has_view() const
{
    return !d_->views_.empty();
}

sheet_view worksheet::view(std::size_t index) const
{
    return d_->views_.at(index);
}

void worksheet::add_view(const sheet_view &new_view)
{
    d_->views_.push_back(new_view);
}

void worksheet::register_comments_in_manifest()
{
    workbook().register_comments_in_manifest(*this);
}

bool worksheet::has_header_footer() const
{
    return d_->header_footer_.is_set();
}

void worksheet::header_footer(const class header_footer &hf)
{
    d_->header_footer_ = hf;
}

void worksheet::page_break_at_row(row_t row)
{
    d_->row_breaks_.push_back(row);
}

const std::vector<row_t> &worksheet::page_break_rows() const
{
    return d_->row_breaks_;
}

void worksheet::page_break_at_column(xlnt::column_t column)
{
    d_->column_breaks_.push_back(column);
}

const std::vector<column_t> &worksheet::page_break_columns() const
{
    return d_->column_breaks_;
}

double worksheet::column_width(column_t column) const
{
    static const auto DefaultColumnWidth = 51.85;

    if (has_column_properties(column))
    {
        return column_properties(column).width.get();
    }
    else
    {
        return points_to_pixels(DefaultColumnWidth, 96.0);
    }
}

double worksheet::row_height(row_t row) const
{
    static const auto DefaultRowHeight = 15.0;

    if (has_row_properties(row))
    {
        return row_properties(row).height.get();
    }
    else
    {
        return points_to_pixels(DefaultRowHeight, 96.0);
    }
}

} // namespace xlnt
