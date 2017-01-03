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
#include <xlnt/worksheet/range_reference.hpp>
#include <xlnt/worksheet/worksheet.hpp>

namespace xlnt {

const_range_iterator::const_range_iterator(const worksheet &ws,
    const range_reference &start_cell, major_order order)
    : ws_(ws.d_),
      current_cell_(start_cell.top_left()),
      range_(start_cell),
      order_(order)
{
}

const_range_iterator::const_range_iterator(const const_range_iterator &other)
{
    *this = other;
}

bool const_range_iterator::operator==(const const_range_iterator &other) const
{
    return ws_ == other.ws_
        && current_cell_ == other.current_cell_
        && order_ == other.order_;
}

bool const_range_iterator::operator!=(const const_range_iterator &other) const
{
    return !(*this == other);
}

const_range_iterator &const_range_iterator::operator--()
{
    if (order_ == major_order::row)
    {
        current_cell_.row(current_cell_.row() - 1);
    }
    else
    {
        current_cell_.column_index(current_cell_.column_index() - 1);
    }

    return *this;
}

const_range_iterator const_range_iterator::operator--(int)
{
    const_range_iterator old = *this;
    --*this;

    return old;
}

const_range_iterator &const_range_iterator::operator++()
{
    if (order_ == major_order::row)
    {
        current_cell_.row(current_cell_.row() + 1);
    }
    else
    {
        current_cell_.column_index(current_cell_.column_index() + 1);
    }

    return *this;
}

const_range_iterator const_range_iterator::operator++(int)
{
    const_range_iterator old = *this;
    ++*this;

    return old;
}

const cell_vector const_range_iterator::operator*() const
{
    if (order_ == major_order::row)
    {
        range_reference reference(range_.top_left().column_index(), current_cell_.row(),
            range_.bottom_right().column_index(), current_cell_.row());
        return cell_vector(ws_, reference, order_);
    }

    range_reference reference(current_cell_.column_index(), range_.top_left().row(),
        current_cell_.column_index(), range_.bottom_right().row());
    return cell_vector(ws_, reference, order_);
}

} // namespace xlnt
