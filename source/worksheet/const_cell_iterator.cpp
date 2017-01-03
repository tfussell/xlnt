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
#include <xlnt/cell/cell_reference.hpp>
#include <xlnt/worksheet/const_cell_iterator.hpp>
#include <xlnt/worksheet/major_order.hpp>

namespace xlnt {

const_cell_iterator::const_cell_iterator(worksheet ws, const cell_reference &start_cell, major_order order)
    : ws_(ws),
      current_cell_(start_cell),
      range_(start_cell.to_range()),
      order_(order)
{
}

const_cell_iterator::const_cell_iterator(const const_cell_iterator &other)
{
    *this = other;
}

bool const_cell_iterator::operator==(const const_cell_iterator &other) const
{
    return ws_ == other.ws_ && current_cell_ == other.current_cell_ && order_ == other.order_;
}

bool const_cell_iterator::operator!=(const const_cell_iterator &other) const
{
    return !(*this == other);
}

const_cell_iterator &const_cell_iterator::operator--()
{
    if (order_ == major_order::row)
    {
        current_cell_.column_index(current_cell_.column_index() - 1);
    }
    else
    {
        current_cell_.row(current_cell_.row() - 1);
    }

    return *this;
}

const_cell_iterator const_cell_iterator::operator--(int)
{
    const_cell_iterator old = *this;
    --*this;

    return old;
}

const_cell_iterator &const_cell_iterator::operator++()
{
    if (order_ == major_order::row)
    {
        current_cell_.column_index(current_cell_.column_index() + 1);
    }
    else
    {
        current_cell_.row(current_cell_.row() + 1);
    }

    return *this;
}

const_cell_iterator const_cell_iterator::operator++(int)
{
    const_cell_iterator old = *this;
    ++*this;

    return old;
}

const cell const_cell_iterator::operator*() const
{
    return ws_.cell(current_cell_);
}

} // namespace xlnt
