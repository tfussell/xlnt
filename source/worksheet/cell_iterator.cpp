// Copyright (c) 2014-2018 Thomas Fussell
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
#include <xlnt/worksheet/cell_iterator.hpp>
#include <xlnt/worksheet/major_order.hpp>

namespace xlnt {

cell_iterator::cell_iterator(worksheet ws, const cell_reference &cursor,
    const range_reference &bounds, major_order order, bool skip_null, bool wrap)
    : ws_(ws),
      cursor_(cursor),
      bounds_(bounds),
      order_(order),
      skip_null_(skip_null),
      wrap_(wrap)
{
    if (skip_null && !ws.has_cell(cursor_))
    {
        (*this)++; // move to the next non-empty cell or one past the end if none exists
    }
}

const_cell_iterator::const_cell_iterator(worksheet ws, const cell_reference &cursor,
    const range_reference &bounds, major_order order, bool skip_null, bool wrap)
    : ws_(ws),
      cursor_(cursor),
      bounds_(bounds),
      order_(order),
      skip_null_(skip_null),
      wrap_(wrap)
{
    if (skip_null && !ws.has_cell(cursor_))
    {
        (*this)++; // move to the next non-empty cell or one past the end if none exists
    }
}

bool cell_iterator::operator==(const cell_iterator &other) const
{
    return ws_ == other.ws_
        && cursor_ == other.cursor_
        && bounds_ == other.bounds_
        && order_ == other.order_
        && skip_null_ == other.skip_null_
        && wrap_ == other.wrap_;
}

bool const_cell_iterator::operator==(const const_cell_iterator &other) const
{
    return ws_ == other.ws_
        && cursor_ == other.cursor_
        && bounds_ == other.bounds_
        && order_ == other.order_
        && skip_null_ == other.skip_null_
        && wrap_ == other.wrap_;
}

bool cell_iterator::operator!=(const cell_iterator &other) const
{
    return !(*this == other);
}

bool const_cell_iterator::operator!=(const const_cell_iterator &other) const
{
    return !(*this == other);
}

cell_iterator &cell_iterator::operator--()
{
    if (order_ == major_order::row)
    {
        if (cursor_.column() > bounds_.top_left().column())
        {
            cursor_.column_index(cursor_.column_index() - 1);
        }

        if (skip_null_)
        {
            while (!ws_.has_cell(cursor_) && cursor_.column() > bounds_.top_left().column())
            {
                cursor_.column_index(cursor_.column_index() - 1);
            }
        }
    }
    else
    {
        if (cursor_.row() > bounds_.top_left().row())
        {
            cursor_.row(cursor_.row() - 1);
        }

        if (skip_null_)
        {
            while (!ws_.has_cell(cursor_) && cursor_.row() > bounds_.top_left().row())
            {
                cursor_.row(cursor_.row() - 1);
            }
        }
    }

    return *this;
}


const_cell_iterator &const_cell_iterator::operator--()
{
    if (order_ == major_order::row)
    {
        if (cursor_.column() > bounds_.top_left().column())
        {
            cursor_.column_index(cursor_.column_index() - 1);
        }

        if (skip_null_)
        {
            while (!ws_.has_cell(cursor_) && cursor_.column() > bounds_.top_left().column())
            {
                cursor_.column_index(cursor_.column_index() - 1);
            }
        }
    }
    else
    {
        if (cursor_.row() > bounds_.top_left().row())
        {
            cursor_.row(cursor_.row() - 1);
        }

        if (skip_null_)
        {
            while (!ws_.has_cell(cursor_) && cursor_.row() > bounds_.top_left().row())
            {
                cursor_.row(cursor_.row() - 1);
            }
        }
    }

    return *this;
}

cell_iterator cell_iterator::operator--(int)
{
    cell_iterator old = *this;
    --*this;

    return old;
}

const_cell_iterator const_cell_iterator::operator--(int)
{
    const_cell_iterator old = *this;
    --*this;

    return old;
}

cell_iterator &cell_iterator::operator++()
{
    if (order_ == major_order::row)
    {
        if (cursor_.column() <= bounds_.bottom_right().column())
        {
            cursor_.column_index(cursor_.column_index() + 1);
        }

        if (skip_null_)
        {
            while (!ws_.has_cell(cursor_) && cursor_.column() <= bounds_.bottom_right().column())
            {
                cursor_.column_index(cursor_.column_index() + 1);
            }
        }
    }
    else
    {
        if (cursor_.row() <= bounds_.bottom_right().row())
        {
            cursor_.row(cursor_.row() + 1);
        }

        if (skip_null_)
        {
            while (!ws_.has_cell(cursor_) && cursor_.row() <= bounds_.bottom_right().row())
            {
                cursor_.row(cursor_.row() + 1);
            }
        }
    }

    return *this;
}

const_cell_iterator &const_cell_iterator::operator++()
{
    if (order_ == major_order::row)
    {
        if (cursor_.column() <= bounds_.bottom_right().column())
        {
            cursor_.column_index(cursor_.column_index() + 1);
        }

        if (skip_null_)
        {
            while (!ws_.has_cell(cursor_) && cursor_.column() <= bounds_.bottom_right().column())
            {
                cursor_.column_index(cursor_.column_index() + 1);
            }
        }
    }
    else
    {
        if (cursor_.row() <= bounds_.bottom_right().row())
        {
            cursor_.row(cursor_.row() + 1);
        }

        if (skip_null_)
        {
            while (!ws_.has_cell(cursor_) && cursor_.row() <= bounds_.bottom_right().row())
            {
                cursor_.row(cursor_.row() + 1);
            }
        }
    }
    
    return *this;
}

cell_iterator cell_iterator::operator++(int)
{
    cell_iterator old = *this;
    ++*this;

    return old;
}

const_cell_iterator const_cell_iterator::operator++(int)
{
    const_cell_iterator old = *this;
    ++*this;

    return old;
}

cell_iterator::reference cell_iterator::operator*()
{
    return ws_.cell(cursor_);
}

const cell_iterator::reference cell_iterator::operator*() const
{
    return ws_.cell(cursor_);
}

const const_cell_iterator::reference const_cell_iterator::operator*() const
{
    return ws_.cell(cursor_);
}
} // namespace xlnt
