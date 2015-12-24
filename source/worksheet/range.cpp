// Copyright (c) 2014-2016 Thomas Fussell
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
#include <xlnt/worksheet/range.hpp>
#include <xlnt/cell/cell.hpp>
#include <xlnt/worksheet/range_reference.hpp>
#include <xlnt/worksheet/worksheet.hpp>

namespace xlnt {

range::range(worksheet ws, const range_reference &reference, major_order order, bool skip_null)
    : ws_(ws), ref_(reference), order_(order), skip_null_(skip_null)
{
}

range::~range()
{
}

cell_vector range::operator[](std::size_t index)
{
    return get_vector(index);
}

range_reference range::get_reference() const
{
    return ref_;
}

std::size_t range::length() const
{
    if (order_ == major_order::row)
    {
        return ref_.get_bottom_right().get_row() - ref_.get_top_left().get_row() + 1;
    }

    return (ref_.get_bottom_right().get_column() - ref_.get_top_left().get_column()).index + 1;
}

bool range::operator==(const range &comparand) const
{
    return ref_ == comparand.ref_ && ws_ == comparand.ws_ && order_ == comparand.order_;
}

cell_vector range::get_vector(std::size_t vector_index)
{
    if (order_ == major_order::row)
    {
        range_reference reference(
            ref_.get_top_left().get_column_index(),
            static_cast<row_t>(static_cast<std::size_t>(ref_.get_top_left().get_row()) + vector_index),
            ref_.get_bottom_right().get_column_index(),
            static_cast<row_t>(static_cast<std::size_t>(ref_.get_top_left().get_row()) + vector_index));

        return cell_vector(ws_, reference, order_);
    }

    range_reference reference(
        static_cast<column_t::index_t>(static_cast<std::size_t>(ref_.get_top_left().get_column().index) + vector_index),
        ref_.get_top_left().get_row(),
        static_cast<column_t::index_t>(static_cast<std::size_t>(ref_.get_top_left().get_column().index) + vector_index),
        ref_.get_bottom_right().get_row());

    return cell_vector(ws_, reference, order_);
}

bool range::contains(const cell_reference &ref)
{
    return ref_.get_top_left().get_column_index() <= ref.get_column_index() &&
           ref_.get_bottom_right().get_column_index() >= ref.get_column_index() &&
           ref_.get_top_left().get_row() <= ref.get_row() && ref_.get_bottom_right().get_row() >= ref.get_row();
}

cell range::get_cell(const cell_reference &ref)
{
    return (*this)[ref.get_row()][ref.get_column().index];
}

range::iterator range::begin()
{
    if (order_ == major_order::row)
    {
        cell_reference top_right(ref_.get_bottom_right().get_column_index(), ref_.get_top_left().get_row());
        range_reference row_range(ref_.get_top_left(), top_right);
        return iterator(ws_, row_range, order_);
    }

    cell_reference bottom_left(ref_.get_top_left().get_column_index(), ref_.get_bottom_right().get_row());
    range_reference row_range(ref_.get_top_left(), bottom_left);
    return iterator(ws_, row_range, order_);
}

range::iterator range::end()
{
    if (order_ == major_order::row)
    {
        auto past_end_row_index = ref_.get_bottom_right().get_row() + 1;
        cell_reference bottom_left(ref_.get_top_left().get_column_index(), past_end_row_index);
        cell_reference bottom_right(ref_.get_bottom_right().get_column_index(), past_end_row_index);
        return iterator(ws_, range_reference(bottom_left, bottom_right), order_);
    }

    auto past_end_column_index = ref_.get_bottom_right().get_column_index() + 1;
    cell_reference top_right(past_end_column_index, ref_.get_top_left().get_row());
    cell_reference bottom_right(past_end_column_index, ref_.get_bottom_right().get_row());
    return iterator(ws_, range_reference(top_right, bottom_right), order_);
}

range::const_iterator range::cbegin() const
{
    if (order_ == major_order::row)
    {
        cell_reference top_right(ref_.get_bottom_right().get_column_index(), ref_.get_top_left().get_row());
        range_reference row_range(ref_.get_top_left(), top_right);
        return const_iterator(ws_, row_range, order_);
    }

    cell_reference bottom_left(ref_.get_top_left().get_column_index(), ref_.get_bottom_right().get_row());
    range_reference row_range(ref_.get_top_left(), bottom_left);
    return const_iterator(ws_, row_range, order_);
}

range::const_iterator range::cend() const
{
    if (order_ == major_order::row)
    {
        auto past_end_row_index = ref_.get_bottom_right().get_row() + 1;
        cell_reference bottom_left(ref_.get_top_left().get_column_index(), past_end_row_index);
        cell_reference bottom_right(ref_.get_bottom_right().get_column_index(), past_end_row_index);
        return const_iterator(ws_, range_reference(bottom_left, bottom_right), order_);
    }

    auto past_end_column_index = ref_.get_bottom_right().get_column_index() + 1;
    cell_reference top_right(past_end_column_index, ref_.get_top_left().get_row());
    cell_reference bottom_right(past_end_column_index, ref_.get_bottom_right().get_row());
    return const_iterator(ws_, range_reference(top_right, bottom_right), order_);
}

cell_vector range::iterator::operator*()
{
    if (order_ == major_order::row)
    {
        range_reference reference(range_.get_top_left().get_column_index(), current_cell_.get_row(),
                                  range_.get_bottom_right().get_column_index(), current_cell_.get_row());
        return cell_vector(ws_, reference, order_);
    }

    range_reference reference(current_cell_.get_column_index(), range_.get_top_left().get_row(),
                              current_cell_.get_column_index(), range_.get_bottom_right().get_row());
    return cell_vector(ws_, reference, order_);
}

range_iterator_2d::range_iterator_2d(worksheet &ws, const range_reference &start_cell, major_order order)
    : ws_(ws.d_), current_cell_(start_cell.get_top_left()), range_(start_cell), order_(order)
{
}

range_iterator_2d::range_iterator_2d(const range_iterator_2d &other)
{
    *this = other;
}

bool range_iterator_2d::operator==(const range_iterator_2d &other) const
{
    return ws_ == other.ws_ && current_cell_ == other.current_cell_ && order_ == other.order_;
}

bool range_iterator_2d::operator!=(const range_iterator_2d &other) const
{
    return !(*this == other);
}

range_iterator_2d &range_iterator_2d::operator--()
{
    if (order_ == major_order::row)
    {
        current_cell_.set_row(current_cell_.get_row() - 1);
    }
    else
    {
        current_cell_.set_column_index(current_cell_.get_column_index() - 1);
    }

    return *this;
}

range_iterator_2d range_iterator_2d::operator--(int)
{
    range_iterator_2d old = *this;
    --*this;
    
    return old;
}

range_iterator_2d &range_iterator_2d::operator++()
{
    if (order_ == major_order::row)
    {
        current_cell_.set_row(current_cell_.get_row() + 1);
    }
    else
    {
        current_cell_.set_column_index(current_cell_.get_column_index() + 1);
    }

    return *this;
}

range_iterator_2d range_iterator_2d::operator++(int)
{
    range_iterator_2d old = *this;
    ++*this;
    
    return old;
}

const_range_iterator_2d::const_range_iterator_2d(const worksheet &ws, const range_reference &start_cell, major_order order)
    : ws_(ws.d_), current_cell_(start_cell.get_top_left()), range_(start_cell), order_(order)
{
}

const_range_iterator_2d::const_range_iterator_2d(const const_range_iterator_2d &other)
{
    *this = other;
}

bool const_range_iterator_2d::operator==(const const_range_iterator_2d &other) const
{
    return ws_ == other.ws_ && current_cell_ == other.current_cell_ && order_ == other.order_;
}

bool const_range_iterator_2d::operator!=(const const_range_iterator_2d &other) const
{
    return !(*this == other);
}

range::const_iterator &range::const_iterator::operator--()
{
    if (order_ == major_order::row)
    {
        current_cell_.set_row(current_cell_.get_row() - 1);
    }
    else
    {
        current_cell_.set_column_index(current_cell_.get_column_index() - 1);
    }

    return *this;
}

range::const_iterator range::const_iterator::operator--(int)
{
    const_range_iterator_2d old = *this;
    --*this;
    
    return old;
}

range::const_iterator &range::const_iterator::operator++()
{
    if (order_ == major_order::row)
    {
        current_cell_.set_row(current_cell_.get_row() + 1);
    }
    else
    {
        current_cell_.set_column_index(current_cell_.get_column_index() + 1);
    }

    return *this;
}

range::const_iterator range::const_iterator::operator++(int)
{
    const_range_iterator_2d old = *this;
    ++*this;
    
    return old;
}

bool range::operator!=(const range &comparand) const
{
    return !(*this == comparand);
}

range::const_iterator range::begin() const
{
    return cbegin();
}

range::const_iterator range::end() const
{
    return cend();
}

} // namespace xlnt
