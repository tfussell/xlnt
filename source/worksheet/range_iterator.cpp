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
#include <xlnt/worksheet/range_iterator.hpp>
#include <xlnt/worksheet/range_reference.hpp>
#include <xlnt/worksheet/worksheet.hpp>

namespace xlnt {

cell_vector range_iterator::operator*() const
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

range_iterator::range_iterator(worksheet &ws, const range_reference &start_cell, major_order order)
    : ws_(ws.d_), current_cell_(start_cell.get_top_left()), range_(start_cell), order_(order)
{
}

range_iterator::range_iterator(const range_iterator &other)
{
    *this = other;
}

bool range_iterator::operator==(const range_iterator &other) const
{
    return ws_ == other.ws_ && current_cell_ == other.current_cell_ && order_ == other.order_;
}

bool range_iterator::operator!=(const range_iterator &other) const
{
    return !(*this == other);
}

range_iterator &range_iterator::operator--()
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

range_iterator range_iterator::operator--(int)
{
    range_iterator old = *this;
    --*this;
    
    return old;
}

range_iterator &range_iterator::operator++()
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

range_iterator range_iterator::operator++(int)
{
    range_iterator old = *this;
    ++*this;
    
    return old;
}

} // namespace xlnt
