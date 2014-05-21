#include "range.h"
#include "cell.h"
#include "range_reference.h"
#include "worksheet.h"

namespace xlnt {

row::iterator::iterator(worksheet ws, const cell_reference &start_cell)
    : ws_(ws),
    current_cell_(start_cell),
    range_(start_cell.to_range())
{
}

row::iterator::~iterator()
{
}

bool row::iterator::operator==(const iterator &rhs)
{
    return ws_ == rhs.ws_ 
        && current_cell_ == rhs.current_cell_;
}

row::iterator row::iterator::operator++(int)
{
    iterator old = *this;
    ++*this;
    return old;
}

row::iterator &row::iterator::operator++()
{
    current_cell_.set_column_index(current_cell_.get_column_index() + 1);
    return *this;
}

cell row::iterator::operator*()
{
    return ws_[current_cell_];
}

row::const_iterator::const_iterator(worksheet ws, const cell_reference &start_cell)
    : ws_(ws),
    current_cell_(start_cell),
    range_(start_cell.to_range())
{
}

row::const_iterator::~const_iterator()
{
}

bool row::const_iterator::operator==(const const_iterator &rhs)
{
    return ws_ == rhs.ws_ 
        && rhs.current_cell_ == current_cell_;
}

row::const_iterator row::const_iterator::operator++(int)
{
    const_iterator old = *this;
    ++*this;
    return old;
}

row::const_iterator &row::const_iterator::operator++()
{
    current_cell_.set_column_index(current_cell_.get_column_index() + 1);
    return *this;
}

const cell row::const_iterator::operator*()
{
    const worksheet ws_const = ws_;
    return ws_const[current_cell_];
}

row::iterator row::begin()
{
    return iterator(ws_, ref_.get_top_left());
}

row::iterator row::end()
{
    auto past_end = ref_.get_bottom_right();
    past_end.set_column_index(past_end.get_column_index() + 1);
    return iterator(ws_, past_end);
}

row::const_iterator row::cbegin() const
{
    return const_iterator(ws_, ref_.get_top_left());
}

row::const_iterator row::cend() const
{
    auto past_end = ref_.get_top_left();
    past_end.set_column_index(past_end.get_column_index() + 1);
    return const_iterator(ws_, past_end);
}

cell row::operator[](std::size_t column_index)
{
    return get_cell(column_index);
}

std::size_t row::num_cells() const
{
    return ref_.get_width() + 1;
}

row::row(worksheet ws, const range_reference &reference)
    : ws_(ws),
    ref_(reference)
{
}

cell row::front()
{
    return get_cell(ref_.get_top_left().get_column_index());
}

cell row::back()
{
    return get_cell(ref_.get_bottom_right().get_column_index());
}

cell row::get_cell(std::size_t column_index)
{
    return ws_.get_cell(ref_.get_top_left().make_offset((int)column_index, 0));
}

range::range(worksheet ws, const range_reference &reference)
    : ws_(ws),
    ref_(reference)
{

}

range::~range()
{
}

row range::operator[](std::size_t row)
{
    return get_row(row);
}

range_reference range::get_reference() const
{
    return ref_;
}

std::size_t range::num_rows() const
{
    return ref_.get_bottom_right().get_row_index() - ref_.get_top_left().get_row_index() + 1;
}

bool range::operator==(const range &comparand) const
{
    return ref_ == comparand.ref_
        && ws_ == comparand.ws_;
}

row range::get_row(std::size_t row_)
{
    range_reference row_reference(ref_.get_top_left().get_column_index(), ref_.get_top_left().get_row_index() + (int)row_,
        ref_.get_bottom_right().get_column_index(), ref_.get_top_left().get_row_index() + (int)row_);
    return row(ws_, row_reference);
}

cell range::get_cell(const cell_reference &ref)
{
    return (*this)[ref.get_row_index()][ref.get_column_index()];
}

range::iterator range::begin()
{
    cell_reference top_right(ref_.get_bottom_right().get_column_index(), 
        ref_.get_top_left().get_row_index());
    range_reference row_range(ref_.get_top_left(), top_right);
    return iterator(ws_, row_range);
}

range::iterator range::end()
{
    auto past_end_row_index = ref_.get_bottom_right().get_row_index() + 1;
    cell_reference bottom_left(ref_.get_top_left().get_column_index(), past_end_row_index);
    cell_reference bottom_right(ref_.get_bottom_right().get_column_index(), past_end_row_index);
    return iterator(ws_, range_reference(bottom_left, bottom_right));
}

range::const_iterator range::cbegin() const
{
    cell_reference top_right(ref_.get_bottom_right().get_column_index(),
        ref_.get_top_left().get_row_index());
    range_reference row_range(ref_.get_top_left(), top_right);
    return const_iterator(ws_, row_range);
}

range::const_iterator range::cend() const
{
    auto past_end_row_index = ref_.get_bottom_right().get_row_index() + 1;
    cell_reference bottom_left(ref_.get_top_left().get_column_index(), past_end_row_index);
    cell_reference bottom_right(ref_.get_bottom_right().get_column_index(), past_end_row_index);
    return const_iterator(ws_, range_reference(bottom_left, bottom_right));
}

range::iterator::iterator(worksheet ws, const range_reference &start_cell)
: ws_(ws),
current_cell_(start_cell.get_top_left()),
range_(start_cell)
{
}

range::iterator::~iterator()
{
}

bool range::iterator::operator==(const iterator &rhs)
{
    return ws_ == rhs.ws_
        && current_cell_ == rhs.current_cell_;
}

range::iterator range::iterator::operator++(int)
{
    iterator old = *this;
    ++*this;
    return old;
}

range::iterator &range::iterator::operator++()
{
    current_cell_.set_row_index(current_cell_.get_row_index() + 1);
    return *this;
}

row range::iterator::operator*()
{
    range_reference row_range(range_.get_top_left().get_column_index(),
        current_cell_.get_row_index(),
        range_.get_bottom_right().get_column_index(),
        current_cell_.get_row_index());
    return row(ws_, row_range);
}

range::const_iterator::const_iterator(worksheet ws, const range_reference &start_cell)
: ws_(ws),
current_cell_(start_cell.get_top_left()),
range_(start_cell)
{
}

range::const_iterator::~const_iterator()
{
}

bool range::const_iterator::operator==(const const_iterator &rhs)
{
    return ws_ == rhs.ws_
        && rhs.current_cell_ == current_cell_;
}

range::const_iterator range::const_iterator::operator++(int)
{
    const_iterator old = *this;
    ++*this;
    return old;
}

range::const_iterator &range::const_iterator::operator++()
{
    current_cell_.set_column_index(current_cell_.get_column_index() + 1);
    return *this;
}

const row range::const_iterator::operator*()
{
    range_reference row_range(range_.get_top_left().get_column_index(),
        current_cell_.get_row_index(),
        range_.get_bottom_right().get_column_index(),
        current_cell_.get_row_index());
    return row(ws_, row_range);
}

} // namespace xlnt
