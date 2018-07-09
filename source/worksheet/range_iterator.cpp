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
#include <xlnt/worksheet/range.hpp>
#include <xlnt/worksheet/range_iterator.hpp>
#include <xlnt/worksheet/range_reference.hpp>
#include <xlnt/worksheet/worksheet.hpp>

namespace xlnt {

range_iterator::reference range_iterator::operator*()
{
    return cell_vector(ws_, cursor_, bounds_, order_, skip_null_, false);
}

const range_iterator::reference range_iterator::operator*() const
{
    return cell_vector(ws_, cursor_, bounds_, order_, skip_null_, false);
}

range_iterator::range_iterator(worksheet &ws, const cell_reference &cursor,
    const range_reference &bounds, major_order order, bool skip_null)
    : ws_(ws),
      cursor_(cursor),
      bounds_(bounds),
      order_(order),
      skip_null_(skip_null)
{
    if (skip_null_ && (**this).empty())
    {
        ++(*this);
    }
}

bool range_iterator::operator==(const range_iterator &other) const
{
    return ws_ == other.ws_
        && cursor_ == other.cursor_
        && order_ == other.order_
        && skip_null_ == other.skip_null_;
}

bool range_iterator::operator!=(const range_iterator &other) const
{
    return !(*this == other);
}

range_iterator &range_iterator::operator--()
{
    if (order_ == major_order::row)
    {
        if (cursor_.row() > bounds_.top_left().row())
        {
            cursor_.row(cursor_.row() - 1);
        }

        if (skip_null_)
        {
            while ((**this).empty() && cursor_.row() > bounds_.top_left().row())
            {
                cursor_.row(cursor_.row() - 1);
            }
        }
    }
    else
    {
        if (cursor_.column() > bounds_.top_left().column())
        {
            cursor_.column_index(cursor_.column_index() - 1);
        }

        if (skip_null_)
        {
            while ((**this).empty() && cursor_.row() > bounds_.top_left().column())
            {
                cursor_.column_index(cursor_.column_index() - 1);
            }
        }
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
        if (cursor_.row() <= bounds_.bottom_right().row())
        {
            cursor_.row(cursor_.row() + 1);
        }
    
        if (skip_null_)
        {
            while ((**this).empty() && cursor_.row() <= bounds_.bottom_right().row())
            {
                cursor_.row(cursor_.row() + 1);
            }
        }
    }
    else
    {
        if (cursor_.column() <= bounds_.bottom_right().column())
        {
            cursor_.column_index(cursor_.column_index() + 1);
        }

        if (skip_null_)
        {
            while ((**this).empty() && cursor_.column() <= bounds_.bottom_right().column())
            {
                cursor_.column_index(cursor_.column_index() + 1);
            }
        }
    }

    return *this;
}

range_iterator range_iterator::operator++(int)
{
    range_iterator old = *this;
    ++*this;

    return old;
}


const_range_iterator::const_range_iterator(const worksheet &ws, const cell_reference &cursor,
    const range_reference &bounds, major_order order, bool skip_null)
    : ws_(ws.d_),
      cursor_(cursor),
      bounds_(bounds),
      order_(order),
      skip_null_(skip_null)
{
    if (skip_null_ && (**this).empty())
    {
        ++(*this);
    }
}

bool const_range_iterator::operator==(const const_range_iterator &other) const
{
    return ws_ == other.ws_
        && cursor_ == other.cursor_
        && order_ == other.order_
        && skip_null_ == other.skip_null_;
}

bool const_range_iterator::operator!=(const const_range_iterator &other) const
{
    return !(*this == other);
}

const_range_iterator &const_range_iterator::operator--()
{
    if (order_ == major_order::row)
    {
        if (cursor_.row() > bounds_.top_left().row())
        {
            cursor_.row(cursor_.row() - 1);
        }

        if (skip_null_)
        {
            while ((**this).empty() && cursor_.row() > bounds_.top_left().row())
            {
                cursor_.row(cursor_.row() - 1);
            }
        }
    }
    else
    {
        if (cursor_.column() > bounds_.top_left().column())
        {
            cursor_.column_index(cursor_.column_index() - 1);
        }

        if (skip_null_)
        {
            while ((**this).empty() && cursor_.row() > bounds_.top_left().column())
            {
                cursor_.column_index(cursor_.column_index() - 1);
            }
        }
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
        if (cursor_.row() <= bounds_.bottom_right().row())
        {
            cursor_.row(cursor_.row() + 1);
        }
    
        if (skip_null_)
        {
            while ((**this).empty() && cursor_.row() <= bounds_.bottom_right().row())
            {
                cursor_.row(cursor_.row() + 1);
            }
        }
    }
    else
    {
        if (cursor_.column() <= bounds_.bottom_right().column())
        {
            cursor_.column_index(cursor_.column_index() + 1);
        }

        if (skip_null_)
        {
            while ((**this).empty() && cursor_.column() <= bounds_.bottom_right().column())
            {
                cursor_.column_index(cursor_.column_index() + 1);
            }
        }
    }

    return *this;
}

const_range_iterator const_range_iterator::operator++(int)
{
    const_range_iterator old = *this;
    ++*this;

    return old;
}

const const_range_iterator::reference const_range_iterator::operator*() const
{
    return cell_vector(ws_, cursor_, bounds_, order_, skip_null_, false);
}

} // namespace xlnt
