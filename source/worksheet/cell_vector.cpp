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
#include <xlnt/worksheet/cell_iterator.hpp>
#include <xlnt/worksheet/cell_vector.hpp>
#include <xlnt/worksheet/const_cell_iterator.hpp>

namespace xlnt {

cell_vector::iterator cell_vector::begin()
{
    return iterator(ws_, ref_.top_left(), ref_, order_);
}

cell_vector::iterator cell_vector::end()
{
    if (order_ == major_order::row)
    {
        auto past_end = ref_.bottom_right();
        past_end.column_index(past_end.column_index() + 1);

        return iterator(ws_, past_end, ref_, order_);
    }

    auto past_end = ref_.bottom_right();
    past_end.row(past_end.row() + 1);

    return iterator(ws_, past_end, ref_, order_);
}

cell_vector::const_iterator cell_vector::cbegin() const
{
    return const_iterator(ws_, ref_.top_left(), order_);
}

cell_vector::const_iterator cell_vector::cend() const
{
    if (order_ == major_order::row)
    {
        auto past_end = ref_.bottom_right();
        past_end.column_index(past_end.column_index() + 1);

        return const_iterator(ws_, past_end, order_);
    }

    auto past_end = ref_.bottom_right();
    past_end.row(past_end.row() + 1);

    return const_iterator(ws_, past_end, order_);
}

cell cell_vector::operator[](std::size_t cell_index)
{
    if (order_ == major_order::row)
    {
        return ws_.cell(ref_.top_left().make_offset(static_cast<int>(cell_index), 0));
    }

    return ws_.cell(ref_.top_left().make_offset(0, static_cast<int>(cell_index)));
}

cell_vector::cell_vector(worksheet ws, const range_reference &reference, major_order order)
    : ws_(ws), ref_(reference), order_(order)
{
}

cell cell_vector::front()
{
    if (order_ == major_order::row)
    {
        return ws_.cell(ref_.top_left());
    }

    return ws_.cell(ref_.top_left());
}

cell cell_vector::back()
{
    if (order_ == major_order::row)
    {
        return ws_.cell(ref_.bottom_right());
    }

    return ws_.cell(ref_.bottom_right());
}

std::size_t cell_vector::length() const
{
    return order_ == major_order::row ? ref_.width() + 1 : ref_.height() + 1;
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

} // namespace xlnt
