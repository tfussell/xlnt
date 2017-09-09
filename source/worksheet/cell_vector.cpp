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

#include <algorithm> // std::all_of

#include <xlnt/cell/cell.hpp>
#include <xlnt/worksheet/cell_iterator.hpp>
#include <xlnt/worksheet/cell_vector.hpp>

namespace xlnt {

cell_vector::cell_vector(worksheet ws, const cell_reference &cursor,
    const range_reference &bounds, major_order order, bool skip_null, bool wrap)
    : ws_(ws),
      cursor_(cursor),
      bounds_(bounds),
      order_(order),
      skip_null_(skip_null),
      wrap_(wrap)
{
}

cell_vector::iterator cell_vector::begin()
{
    return iterator(ws_, cursor_, bounds_, order_, skip_null_, wrap_);
}

cell_vector::iterator cell_vector::end()
{
    auto past_end = cursor_;
    
    if (order_ == major_order::row)
    {
        past_end.column_index(bounds_.bottom_right().column_index() + 1);
    }
    else
    {
        past_end.row(bounds_.bottom_right().row() + 1);
    }

    return iterator(ws_, past_end, bounds_, order_, skip_null_, wrap_);
}

cell_vector::const_iterator cell_vector::cbegin() const
{
    return const_iterator(ws_, cursor_, bounds_, order_, skip_null_, wrap_);
}

cell_vector::const_iterator cell_vector::cend() const
{
    auto past_end = cursor_;
    
    if (order_ == major_order::row)
    {
        past_end.column_index(bounds_.bottom_right().column_index() + 1);
    }
    else
    {
        past_end.row(bounds_.bottom_right().row() + 1);
    }

    return const_iterator(ws_, past_end, bounds_, order_, skip_null_, wrap_);
}

bool cell_vector::empty() const
{
    return begin() == end();
}

cell cell_vector::front()
{
    return *begin();
}

const cell cell_vector::front() const
{
    return *cbegin();
}

cell cell_vector::back()
{
    return *(--end());
}

const cell cell_vector::back() const
{
    return *(--cend());
}

std::size_t cell_vector::length() const
{
    return order_ == major_order::row ? bounds_.width() : bounds_.height();
}

cell_vector::const_iterator cell_vector::begin() const
{
    return cbegin();
}
cell_vector::const_iterator cell_vector::end() const
{
    return cend();
}

cell_vector::reverse_iterator cell_vector::rbegin()
{
    return reverse_iterator(end());
}

cell_vector::reverse_iterator cell_vector::rend()
{
    return reverse_iterator(begin());
}

cell_vector::const_reverse_iterator cell_vector::crbegin() const
{
    return const_reverse_iterator(cend());
}

cell_vector::const_reverse_iterator cell_vector::rbegin() const
{
    return crbegin();
}

cell_vector::const_reverse_iterator cell_vector::crend() const
{
    return const_reverse_iterator(cbegin());
}

cell_vector::const_reverse_iterator cell_vector::rend() const
{
    return crend();
}

cell cell_vector::operator[](std::size_t cell_index)
{
    if (order_ == major_order::row)
    {
        return ws_.cell(cursor_.make_offset(static_cast<int>(cell_index), 0));
    }

    return ws_.cell(cursor_.make_offset(0, static_cast<int>(cell_index)));
}

} // namespace xlnt
