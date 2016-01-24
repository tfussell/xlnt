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
#pragma once

#include <xlnt/xlnt_config.hpp>
#include <xlnt/cell/cell_reference.hpp>
#include <xlnt/worksheet/major_order.hpp>
#include <xlnt/worksheet/range_reference.hpp>

namespace xlnt {

class cell_vector;
class worksheet;

namespace detail {
struct worksheet_impl;
}

/// <summary>
/// An iterator used by worksheet and range for traversing
/// a 2D grid of cells by row/column then across that row/column.
/// </summary>
class XLNT_CLASS range_iterator : public std::iterator<std::bidirectional_iterator_tag, cell_vector>
{
public:
    range_iterator(worksheet &ws, const range_reference &start_cell, major_order order = major_order::row);

    range_iterator(const range_iterator &other);

    cell_vector operator*();

    bool operator==(const range_iterator &other) const;

    bool operator!=(const range_iterator &other) const;

    range_iterator &operator--();

    range_iterator operator--(int);

    range_iterator &operator++();

    range_iterator operator++(int);

private:
    detail::worksheet_impl *ws_;
    cell_reference current_cell_;
    range_reference range_;
    major_order order_;
};

} // namespace xlnt
