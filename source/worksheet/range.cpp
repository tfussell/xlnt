// Copyright (c) 2014-2017 Thomas Fussell
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
#include <xlnt/cell/cell.hpp>
#include <xlnt/worksheet/const_range_iterator.hpp>
#include <xlnt/worksheet/range.hpp>
#include <xlnt/worksheet/range_iterator.hpp>
#include <xlnt/worksheet/range_reference.hpp>
#include <xlnt/worksheet/worksheet.hpp>

namespace xlnt {

range::range(worksheet ws, const range_reference &reference, major_order order, bool skip_null)
    : ws_(ws),
      ref_(reference),
      order_(order),
      skip_null_(skip_null)
{
}

range::~range()
{
}

cell_vector range::operator[](std::size_t index)
{
    return vector(index);
}

range_reference range::reference() const
{
    return ref_;
}

std::size_t range::length() const
{
    if (order_ == major_order::row)
    {
        return ref_.bottom_right().row() - ref_.top_left().row() + 1;
    }

    return (ref_.bottom_right().column() - ref_.top_left().column()).index + 1;
}

bool range::operator==(const range &comparand) const
{
    return ref_ == comparand.ref_
        && ws_ == comparand.ws_
        && order_ == comparand.order_;
}

cell_vector range::vector(std::size_t vector_index)
{
    range_reference vector_ref = ref_;

    if (order_ == major_order::row)
    {
        auto row = ref_.top_left().row() + static_cast<row_t>(vector_index);
        vector_ref.top_left().row(row);
        vector_ref.bottom_right().row(row);
    }
    else
    {
        auto column = ref_.top_left().column() + static_cast<column_t::index_t>(vector_index);
        vector_ref.top_left().column_index(column);
        vector_ref.bottom_right().column_index(column);
    }

    return cell_vector(ws_, vector_ref, order_);
}

bool range::contains(const cell_reference &ref)
{
    return ref_.top_left().column_index() <= ref.column_index()
        && ref_.bottom_right().column_index() >= ref.column_index()
        && ref_.top_left().row() <= ref.row()
        && ref_.bottom_right().row() >= ref.row();
}

cell range::cell(const cell_reference &ref)
{
    return (*this)[ref.row() - 1][ref.column().index - 1];
}

range::iterator range::begin()
{
    if (order_ == major_order::row)
    {
        cell_reference top_right(ref_.bottom_right().column_index(), ref_.top_left().row());
        range_reference row_range(ref_.top_left(), top_right);
        return iterator(ws_, row_range, ref_, order_);
    }

    cell_reference bottom_left(ref_.top_left().column_index(), ref_.bottom_right().row());
    range_reference row_range(ref_.top_left(), bottom_left);
    return iterator(ws_, row_range, ref_, order_);
}

range::iterator range::end()
{
    if (order_ == major_order::row)
    {
        auto past_end_row_index = ref_.bottom_right().row() + 1;
        cell_reference bottom_left(ref_.top_left().column_index(), past_end_row_index);
        cell_reference bottom_right(ref_.bottom_right().column_index(), past_end_row_index);
        return iterator(ws_, range_reference(bottom_left, bottom_right), ref_, order_);
    }

    auto past_end_column_index = ref_.bottom_right().column_index() + 1;
    cell_reference top_right(past_end_column_index, ref_.top_left().row());
    cell_reference bottom_right(past_end_column_index, ref_.bottom_right().row());
    return iterator(ws_, range_reference(top_right, bottom_right), ref_, order_);
}

range::const_iterator range::cbegin() const
{
    if (order_ == major_order::row)
    {
        cell_reference top_right(ref_.bottom_right().column_index(), ref_.top_left().row());
        range_reference row_range(ref_.top_left(), top_right);
        return const_iterator(ws_, row_range, order_);
    }

    cell_reference bottom_left(ref_.top_left().column_index(), ref_.bottom_right().row());
    range_reference row_range(ref_.top_left(), bottom_left);
    return const_iterator(ws_, row_range, order_);
}

range::const_iterator range::cend() const
{
    if (order_ == major_order::row)
    {
        auto past_end_row_index = ref_.bottom_right().row() + 1;
        cell_reference bottom_left(ref_.top_left().column_index(), past_end_row_index);
        cell_reference bottom_right(ref_.bottom_right().column_index(), past_end_row_index);
        return const_iterator(ws_, range_reference(bottom_left, bottom_right), order_);
    }

    auto past_end_column_index = ref_.bottom_right().column_index() + 1;
    cell_reference top_right(past_end_column_index, ref_.top_left().row());
    cell_reference bottom_right(past_end_column_index, ref_.bottom_right().row());
    return const_iterator(ws_, range_reference(top_right, bottom_right), order_);
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

range::reverse_iterator range::rbegin()
{
    return reverse_iterator(end());
}

range::reverse_iterator range::rend()
{
    return reverse_iterator(begin());
}

range::const_reverse_iterator range::crbegin() const
{
    return const_reverse_iterator(cend());
}

range::const_reverse_iterator range::rbegin() const
{
    return crbegin();
}

range::const_reverse_iterator range::crend() const
{
    return const_reverse_iterator(cbegin());
}

range::const_reverse_iterator range::rend() const
{
    return crend();
}

} // namespace xlnt
