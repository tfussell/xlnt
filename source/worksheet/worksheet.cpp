// Copyright (c) 2014-2021 Thomas Fussell
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
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
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
#include <xlnt/utils/numeric.hpp>
#include <xlnt/workbook/named_range.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/workbook/worksheet_iterator.hpp>
#include <xlnt/worksheet/cell_iterator.hpp>
#include <xlnt/worksheet/column_properties.hpp>
#include <xlnt/worksheet/header_footer.hpp>
#include <xlnt/worksheet/range.hpp>
#include <xlnt/worksheet/range_iterator.hpp>
#include <xlnt/worksheet/range_reference.hpp>
#include <xlnt/worksheet/row_properties.hpp>
#include <xlnt/worksheet/worksheet.hpp>
#include <detail/constants.hpp>
#include <detail/default_case.hpp>
#include <detail/implementations/cell_impl.hpp>
#include <detail/implementations/workbook_impl.hpp>
#include <detail/implementations/worksheet_impl.hpp>
#include <detail/unicode.hpp>

namespace {

int points_to_pixels(double points, double dpi)
{
    return static_cast<int>(std::ceil(points * dpi / 72));
}

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
    catch (xlnt::invalid_cell_reference &)
    {
        // name is not a valid reference, that's good
    }

    std::vector<named_range::target> targets;
    targets.push_back({*this, reference});

    d_->named_ranges_[name] = xlnt::named_range(name, targets);
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
    auto cell_iter = d_->cell_map_.begin();

    while (cell_iter != d_->cell_map_.end())
    {
        if (xlnt::cell(&cell_iter->second).garbage_collectible())
        {
            cell_iter = d_->cell_map_.erase(cell_iter);
        }
        else
        {
            ++cell_iter;
        }
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
    // do no work if we don't need to
    if (d_->title_ == title)
    {
        return;
    }
    // excel limits worksheet titles to 31 characters
    if (title.empty() || detail::string_length(title) > 31)
    {
        throw invalid_sheet_title(title);
    }
    // invalid characters in a worksheet name
    if (title.find_first_of("*:/\\?[]") != std::string::npos)
    {
        throw invalid_sheet_title(title);
    }
    // try and insert the new name into the worksheets map
    // if the insert fails, we have a duplicate sheet name
    auto insert_result = workbook().d_->sheet_title_rel_id_map_.insert(
        std::make_pair(title, workbook().d_->sheet_title_rel_id_map_[d_->title_]));
    if (!insert_result.second) // insert failed, duplication detected
    {
        throw invalid_sheet_title(title);
    }
    // if the insert succeeded (i.e. wasn't a duplicate sheet name)
    // update the worksheet title and remove the old relation
    workbook().d_->sheet_title_rel_id_map_.erase(d_->title_);
    d_->title_ = title;

    workbook().update_sheet_properties();
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
    if (ref == "A1")
    {
        unfreeze_panes();
        return;
    }
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
    if (ref.column() == "A") // no column is frozen
    {
        primary_view.add_selection(selection(pane_corner::bottom_left, ref));
        primary_view.pane().active_pane = pane_corner::bottom_left;
        primary_view.pane().y_split = ref.row() - 1;
    }
    else if (ref.row() == 1) // no row is frozen
    {
        primary_view.add_selection(selection(pane_corner::top_right, ref));
        primary_view.pane().active_pane = pane_corner::top_right;
        primary_view.pane().x_split = ref.column_index() - 1;
    }
    else // column and row is frozen
    {
        primary_view.add_selection(selection(pane_corner::top_right, cell_reference(ref.column(), 1)));
        primary_view.add_selection(selection(pane_corner::bottom_left, cell_reference(1, ref.row())));
        primary_view.add_selection(selection(pane_corner::bottom_right, ref));
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

void worksheet::active_cell(const cell_reference &ref)
{
    if (!has_view())
    {
        d_->views_.push_back(sheet_view());
    }

    auto &primary_view = d_->views_.front();

    if (!primary_view.has_selections())
    {
        primary_view.add_selection(selection(pane_corner::bottom_right, ref));
    }
    else
    {
        primary_view.selection(0).active_cell(ref);
    }
}

bool worksheet::has_active_cell() const
{
    if (!has_view()) return false;
    auto &primary_view = d_->views_.front();
    if (!primary_view.has_selections()) return false;
    auto primary_selection = primary_view.selection(0);

    return primary_selection.has_active_cell();
}

cell_reference worksheet::active_cell() const
{
    if (!has_view())
    {
        throw xlnt::exception("Worksheet has no view.");
    }

    auto &primary_view = d_->views_.front();

    if (!primary_view.has_selections())
    {
        throw xlnt::exception("Default worksheet view has no selections.");
    }

    return primary_view.selection(0).active_cell();
}

cell worksheet::cell(const cell_reference &reference)
{
    auto match = d_->cell_map_.find(reference);
    if (match == d_->cell_map_.end())
    {
        auto impl = detail::cell_impl();
        impl.parent_ = d_;
        impl.column_ = reference.column_index();
        impl.row_ = reference.row();

        match = d_->cell_map_.emplace(reference, impl).first;
    }
    return xlnt::cell(&match->second);
}

const cell worksheet::cell(const cell_reference &reference) const
{
    return xlnt::cell(&d_->cell_map_.at(reference));
}

cell worksheet::cell(xlnt::column_t column, row_t row)
{
    return cell(cell_reference(column, row));
}

const cell worksheet::cell(xlnt::column_t column, row_t row) const
{
    return cell(cell_reference(column, row));
}

bool worksheet::has_cell(const cell_reference &reference) const
{
    const auto cell = d_->cell_map_.find(reference);
    return cell != d_->cell_map_.cend();
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

const range worksheet::named_range(const std::string &name) const
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

    auto lowest = constants::max_column();

    for (auto &cell : d_->cell_map_)
    {
        lowest = std::min(lowest, cell.first.column());
    }

    return lowest;
}

column_t worksheet::lowest_column_or_props() const
{
    auto lowest = lowest_column();

    if (d_->cell_map_.empty() && !d_->column_properties_.empty())
    {
        lowest = d_->column_properties_.begin()->first;
    }

    for (auto &props : d_->column_properties_)
    {
        lowest = std::min(lowest, props.first);
    }

    return lowest;
}

row_t worksheet::lowest_row() const
{
    if (d_->cell_map_.empty())
    {
        return constants::min_row();
    }

    auto lowest = constants::max_row();

    for (auto &cell : d_->cell_map_)
    {
        lowest = std::min(lowest, cell.first.row());
    }

    return lowest;
}

row_t worksheet::lowest_row_or_props() const
{
    auto lowest = lowest_row();

    if (d_->cell_map_.empty() && !d_->row_properties_.empty())
    {
        lowest = d_->row_properties_.begin()->first;
    }

    for (auto &props : d_->row_properties_)
    {
        lowest = std::min(lowest, props.first);
    }

    return lowest;
}

row_t worksheet::highest_row() const
{
    auto highest = constants::min_row();

    for (auto &cell : d_->cell_map_)
    {
        highest = std::max(highest, cell.first.row());
    }

    return highest;
}

row_t worksheet::highest_row_or_props() const
{
    auto highest = highest_row();

    if (d_->cell_map_.empty() && !d_->row_properties_.empty())
    {
        highest = d_->row_properties_.begin()->first;
    }

    for (auto &props : d_->row_properties_)
    {
        highest = std::max(highest, props.first);
    }

    return highest;
}

column_t worksheet::highest_column() const
{
    auto highest = constants::min_column();

    for (auto &cell : d_->cell_map_)
    {
        highest = std::max(highest, cell.first.column());
    }

    return highest;
}

column_t worksheet::highest_column_or_props() const
{
    auto highest = highest_column();

    if (d_->cell_map_.empty() && !d_->column_properties_.empty())
    {
        highest = d_->column_properties_.begin()->first;
    }

    for (auto &props : d_->column_properties_)
    {
        highest = std::max(highest, props.first);
    }

    return highest;
}

range_reference worksheet::calculate_dimension(bool skip_null) const
{
    // partially optimised version of:
    // return range_reference(lowest_column(), lowest_row_or_props(),
    //                        highest_column(), highest_row_or_props());
    //
    if (d_->cell_map_.empty() && d_->row_properties_.empty())
    {
        return range_reference(constants::min_column(), constants::min_row(),
            constants::min_column(), constants::min_row());
    }

    // if skip_null = false, min row = min_row() and min column = min_column()
    // in order to include first empty rows and columns
    row_t min_row_prop = skip_null? constants::max_row() : constants::min_row();
    row_t max_row_prop = constants::min_row();
    for (const auto &row_prop : d_->row_properties_)
    {
        if(skip_null){
            min_row_prop = std::min(min_row_prop, row_prop.first);
        }
        max_row_prop = std::max(max_row_prop, row_prop.first);
    }
    if (d_->cell_map_.empty())
    {
        return range_reference(constants::min_column(), min_row_prop,
            constants::min_column(), max_row_prop);
    }
    // find min and max row/column in cell map
    column_t min_col = skip_null? constants::max_column() : constants::min_column();
    column_t max_col = constants::min_column();
    row_t min_row = min_row_prop;
    row_t max_row = max_row_prop;
    for (auto &c : d_->cell_map_)
    {
        if(skip_null){
            min_col = std::min(min_col, c.second.column_);
            min_row = std::min(min_row, c.second.row_);
        }
        max_col = std::max(max_col, c.second.column_);
        max_row = std::max(max_row, c.second.row_);
    }
    return range_reference(min_col, min_row, max_col, max_row);
}

range worksheet::range(const std::string &reference_string)
{
    if (has_named_range(reference_string))
    {
        return named_range(reference_string);
    }

    return range(range_reference(reference_string));
}

const range worksheet::range(const std::string &reference_string) const
{
    if (has_named_range(reference_string))
    {
        return named_range(reference_string);
    }

    return range(range_reference(reference_string));
}

range worksheet::range(const range_reference &reference)
{
    return xlnt::range(*this, reference);
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
                if (cell.data_type() == cell::type::shared_string)
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

row_t worksheet::next_row() const
{
    auto row = highest_row() + 1;

    if (row == 2 && d_->cell_map_.size() == 0)
    {
        row = 1;
    }

    return row;
}

xlnt::range worksheet::rows(bool skip_null)
{
    return xlnt::range(*this, calculate_dimension(skip_null), major_order::row, skip_null);
}

const xlnt::range worksheet::rows(bool skip_null) const
{
    return xlnt::range(*this, calculate_dimension(skip_null), major_order::row, skip_null);
}

xlnt::range worksheet::columns(bool skip_null)
{
    return xlnt::range(*this, calculate_dimension(skip_null), major_order::column, skip_null);
}

const xlnt::range worksheet::columns(bool skip_null) const
{
    return xlnt::range(*this, calculate_dimension(skip_null), major_order::column, skip_null);
}

/*
//TODO: finish implementing cell_iterator wrapping before uncommenting

cell_vector worksheet::cells(bool skip_null)
{
    const auto dimension = calculate_dimension();
    return cell_vector(*this, dimension.top_left(), dimension, major_order::row, skip_null, true);
}

const cell_vector worksheet::cells(bool skip_null) const
{
    const auto dimension = calculate_dimension();
    return cell_vector(*this, dimension.top_left(), dimension, major_order::row, skip_null, true);
}
*/

void worksheet::clear_cell(const cell_reference &ref)
{
    d_->cell_map_.erase(ref);
    // TODO: garbage collect newly unreferenced resources such as styles?
}

void worksheet::clear_row(row_t row)
{
    for (auto it = d_->cell_map_.begin(); it != d_->cell_map_.end();)
    {
        if (it->first.row() == row)
        {
            it = d_->cell_map_.erase(it);
        }
        else
        {
            ++it;
        }
    }
    d_->row_properties_.erase(row);
    // TODO: garbage collect newly unreferenced resources such as styles?
}

void worksheet::insert_rows(row_t row, std::uint32_t amount)
{
    move_cells(row, amount, row_or_col_t::row);
}

void worksheet::insert_columns(column_t column, std::uint32_t amount)
{
    move_cells(column.index, amount, row_or_col_t::column);
}

void worksheet::delete_rows(row_t row, std::uint32_t amount)
{
    move_cells(row + amount, amount, row_or_col_t::row, true);
}

void worksheet::delete_columns(column_t column, std::uint32_t amount)
{
    move_cells(column.index + amount, amount, row_or_col_t::column, true);
}

void worksheet::move_cells(std::uint32_t min_index, std::uint32_t amount, row_or_col_t row_or_col, bool reverse)
{
    if (reverse && amount > min_index)
    {
        throw xlnt::invalid_parameter();
    }

    if ((!reverse && row_or_col == row_or_col_t::row && min_index > constants::max_row() - amount) || (!reverse && row_or_col == row_or_col_t::column && min_index > constants::max_column() - amount))
    {
        throw xlnt::exception("Cannot move cells as they would be outside the maximum bounds of the spreadsheet");
    }

    std::vector<detail::cell_impl> cells_to_move;

    auto cell_iter = d_->cell_map_.cbegin();
    while (cell_iter != d_->cell_map_.cend())
    {
        std::uint32_t current_index;
        switch (row_or_col)
        {
        case row_or_col_t::row:
            current_index = cell_iter->first.row();
            break;
        case row_or_col_t::column:
            current_index = cell_iter->first.column().index;
            break;
        default:
            throw xlnt::unhandled_switch_case();
        }

        if (current_index >= min_index) // extract cells to be moved
        {
            auto cell = cell_iter->second;
            if (row_or_col == row_or_col_t::row)
            {
                cell.row_ = reverse ? cell.row_ - amount : cell.row_ + amount;
            }
            else if (row_or_col == row_or_col_t::column)
            {
                cell.column_ = reverse ? cell.column_.index - amount : cell.column_.index + amount;
            }

            cells_to_move.push_back(cell);
            cell_iter = d_->cell_map_.erase(cell_iter);
        }
        else if (reverse && current_index >= min_index - amount) // delete destination cells
        {
            cell_iter = d_->cell_map_.erase(cell_iter);
        }
        else // skip other cells
        {
            ++cell_iter;
        }
    }

    for (auto &cell : cells_to_move)
    {
        d_->cell_map_[cell_reference(cell.column_, cell.row_)] = cell;
    }

    if (row_or_col == row_or_col_t::row)
    {
        std::vector<std::pair<row_t, xlnt::row_properties>> properties_to_move;

        auto row_prop_iter = d_->row_properties_.cbegin();
        while (row_prop_iter != d_->row_properties_.cend())
        {
            auto current_row = row_prop_iter->first;
            if (current_row >= min_index) // extract properties that need to be moved
            {
                auto tmp_row = reverse ? current_row - amount : current_row + amount;
                properties_to_move.push_back({tmp_row, row_prop_iter->second});
                row_prop_iter = d_->row_properties_.erase(row_prop_iter);
            }
            else if (reverse && current_row >= min_index - amount) // clear properties of destination when in reverse
            {
                row_prop_iter = d_->row_properties_.erase(row_prop_iter);
            }
            else // skip the rest
            {
                ++row_prop_iter;
            }
        }

        for (const auto &prop : properties_to_move)
        {
            add_row_properties(prop.first, prop.second);
        }
    }
    else if (row_or_col == row_or_col_t::column)
    {
        std::vector<std::pair<column_t, xlnt::column_properties>> properties_to_move;

        auto col_prop_iter = d_->column_properties_.cbegin();
        while (col_prop_iter != d_->column_properties_.cend())
        {
            auto current_col = col_prop_iter->first.index;
            if (current_col >= min_index) // extract properties that need to be moved
            {
                auto tmp_column = column_t(reverse ? current_col - amount : current_col + amount);
                properties_to_move.push_back({tmp_column, col_prop_iter->second});
                col_prop_iter = d_->column_properties_.erase(col_prop_iter);
            }
            else if (reverse && current_col >= min_index - amount) // clear properties of destination when in reverse
            {
                col_prop_iter = d_->column_properties_.erase(col_prop_iter);
            }
            else // skip the rest
            {
                ++col_prop_iter;
            }
        }

        for (auto &prop : properties_to_move)
        {
            add_column_properties(prop.first, prop.second);
        }
    }

    // adjust merged cells
    auto shift_reference = [min_index, amount, row_or_col, reverse](cell_reference &ref) {
        auto index = row_or_col == row_or_col_t::row ? ref.row() : ref.column_index();
        if (index >= min_index)
        {
            auto new_index = reverse ? index - amount : index + amount;
            if (row_or_col == row_or_col_t::row)
            {
                ref.row(new_index);
            }
            else if (row_or_col == row_or_col_t::column)
            {
                ref.column_index(new_index);
            }
        }
    };

    for (auto merged_cell = d_->merged_cells_.begin(); merged_cell != d_->merged_cells_.end(); ++merged_cell)
    {
        cell_reference new_top_left = merged_cell->top_left();
        shift_reference(new_top_left);

        cell_reference new_bottom_right = merged_cell->bottom_right();
        shift_reference(new_bottom_right);

        range_reference new_range{new_top_left, new_bottom_right};
        if (*merged_cell != new_range)
        {
            *merged_cell = new_range;
        }
    }
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

    for (auto &cell : d_->cell_map_)
    {
        if (other.d_->cell_map_.find(cell.first) == other.d_->cell_map_.end())
        {
            return false;
        }

        xlnt::cell this_cell(&cell.second);
        xlnt::cell other_cell(&other.d_->cell_map_[cell.first]);

        if (this_cell.data_type() != other_cell.data_type())
        {
            return false;
        }

        if (this_cell.data_type() == xlnt::cell::type::number
            && !detail::float_equals(this_cell.value<double>(), other_cell.value<double>()))
        {
            return false;
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

bool worksheet::has_named_range(const std::string &name) const
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

cell_reference worksheet::point_pos(int left, int top) const
{
    column_t current_column = 1;
    row_t current_row = 1;

    double left_pos = 0;
    double top_pos = 0;

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

void worksheet::add_row_properties(row_t row, const xlnt::row_properties &props)
{
    d_->row_properties_[row] = props;
}

worksheet::iterator worksheet::begin()
{
    return rows().begin();
}

worksheet::iterator worksheet::end()
{
    return rows().end();
}

worksheet::const_iterator worksheet::cbegin() const
{
    return rows().cbegin();
}

worksheet::const_iterator worksheet::cend() const
{
    return rows().cend();
}

worksheet::const_iterator worksheet::begin() const
{
    return cbegin();
}

worksheet::const_iterator worksheet::end() const
{
    return cend();
}

void worksheet::print_title_rows(row_t start, row_t end)
{
    d_->print_title_rows_ = std::make_pair(start, end);
}

optional<std::pair<row_t, row_t>> worksheet::print_title_rows() const
{
    return d_->print_title_rows_;
}

void worksheet::print_title_cols(column_t start, column_t end)
{
    d_->print_title_cols_ = std::make_pair(start, end);
}

optional<std::pair<column_t, column_t>> worksheet::print_title_cols() const
{
    return d_->print_title_cols_;
}

bool worksheet::has_print_titles() const
{
    return d_->print_title_cols_.is_set() || d_->print_title_rows_.is_set();
}

void worksheet::clear_print_titles()
{
    d_->print_title_rows_.clear();
    d_->print_title_cols_.clear();
}

void worksheet::print_area(const std::string &print_area)
{
    d_->print_area_ = range_reference::make_absolute(range_reference(print_area));
}

range_reference worksheet::print_area() const
{
    return d_->print_area_.get();
}

bool worksheet::has_print_area() const
{
    return d_->print_area_.is_set();
}

void worksheet::clear_print_area()
{
    return d_->print_area_.clear();
}

bool worksheet::has_view() const
{
    return !d_->views_.empty();
}

sheet_view &worksheet::view(std::size_t index) const
{
    return d_->views_.at(index);
}

void worksheet::add_view(const sheet_view &new_view)
{
    d_->views_.push_back(new_view);
}

void worksheet::register_comments_in_manifest()
{
    workbook().register_worksheet_part(*this, relationship_type::comments);
}

void worksheet::register_calc_chain_in_manifest()
{
    workbook().register_workbook_part(relationship_type::calculation_chain);
}

bool worksheet::has_phonetic_properties() const
{
    return d_->phonetic_properties_.is_set();
}

const phonetic_pr &worksheet::phonetic_properties() const
{
    return d_->phonetic_properties_.get();
}

void worksheet::phonetic_properties(const phonetic_pr &phonetic_props)
{
    d_->phonetic_properties_.set(phonetic_props);
}

bool worksheet::has_header_footer() const
{
    return d_->header_footer_.is_set();
}

void worksheet::header_footer(const class header_footer &hf)
{
    d_->header_footer_ = hf;
}

void worksheet::clear_page_breaks()
{
    d_->row_breaks_.clear();
    d_->column_breaks_.clear();
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

    if (has_row_properties(row) && row_properties(row).height.is_set())
    {
        return row_properties(row).height.get();
    }
    else
    {
        return points_to_pixels(DefaultRowHeight, 96.0);
    }
}

void worksheet::garbage_collect_formulae()
{
    workbook().garbage_collect_formulae();
}

void worksheet::parent(xlnt::workbook &wb)
{
    d_->parent_ = &wb;
}

conditional_format worksheet::conditional_format(const range_reference &ref, const condition &when)
{
    return workbook().d_->stylesheet_.get().add_conditional_format_rule(d_, ref, when);
}

path worksheet::path() const
{
    auto rel = referring_relationship();
    return xlnt::path(rel.source().path().parent().append(rel.target().path()));
}

relationship worksheet::referring_relationship() const
{
    auto &manifest = workbook().manifest();
    auto wb_rel = manifest.relationship(xlnt::path("/"),
        relationship_type::office_document);
    auto ws_rel = manifest.relationship(wb_rel.target().path(),
        workbook().d_->sheet_title_rel_id_map_.at(title()));
    return ws_rel;
}

sheet_format_properties worksheet::format_properties() const
{
    return d_->format_properties_;
}

void worksheet::format_properties(const sheet_format_properties &properties)
{
    d_->format_properties_ = properties;
}

bool worksheet::has_drawing() const
{
    return d_->drawing_.is_set();
}

bool worksheet::is_empty() const
{
    return d_->cell_map_.empty();
}

} // namespace xlnt
