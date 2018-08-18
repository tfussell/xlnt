// Copyright (c) 2014-2018 Thomas Fussell
// Copyright (c) 2010-2015 openpyxl
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

#include <cstddef> // std::ptrdiff_t
#include <iterator>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/cell/cell_reference.hpp>
#include <xlnt/worksheet/range_reference.hpp>
#include <xlnt/worksheet/major_order.hpp>
#include <xlnt/worksheet/worksheet.hpp>

namespace xlnt {

enum class major_order;

class cell;
class cell_reference;
class range_reference;

/// <summary>
/// A cell iterator iterates over a 1D range by row or by column.
/// </summary>
class XLNT_API cell_iterator
{
public:
    /// <summary>
    /// iterator tags required for use with standard algorithms and adapters
    /// </summary>
    using iterator_category = std::bidirectional_iterator_tag;
    using value_type = cell;
    using difference_type = std::ptrdiff_t;
    using pointer = cell *;
    using reference = cell; // intentionally value

    /// <summary>
    /// Default constructs a cell_iterator
    /// </summary>
    cell_iterator() = default;

    /// <summary>
    /// Constructs a cell_iterator from a worksheet, range, and iteration settings.
    /// </summary>
    cell_iterator(worksheet ws, const cell_reference &start_cell,
        const range_reference &limits, major_order order, bool skip_null, bool wrap);

    /// <summary>
    /// Constructs a cell_iterator as a copy of an existing cell_iterator.
    /// </summary>
    cell_iterator(const cell_iterator &) = default;

    /// <summary>
    /// Assigns this iterator to match the data in rhs.
    /// </summary>
    cell_iterator &operator=(const cell_iterator &) = default;

    /// <summary>
    /// Constructs a cell_iterator by moving from a cell_iterator temporary
    /// </summary>
    cell_iterator(cell_iterator &&) = default;

    /// <summary>
    /// Assigns this iterator to from a cell_iterator temporary
    /// </summary>
    cell_iterator &operator=(cell_iterator &&) = default;

    /// <summary>
    /// destructor for const_cell_iterator
    /// </summary>
    ~cell_iterator() = default;

    /// <summary>
    /// Dereferences this iterator to return the cell it points to.
    /// </summary>
    reference operator*();

    /// <summary>
    /// Dereferences this iterator to return the cell it points to.
    /// </summary>
    const reference operator*() const;


    /// <summary>
    /// Returns true if this iterator is equivalent to other.
    /// </summary>
    bool operator==(const cell_iterator &other) const;

    /// <summary>
    /// Returns true if this iterator isn't equivalent to other.
    /// </summary>
    bool operator!=(const cell_iterator &other) const;

    /// <summary>
    /// Pre-decrements the iterator to point to the previous cell and
    /// returns a reference to the iterator.
    /// </summary>
    cell_iterator &operator--();

    /// <summary>
    /// Post-decrements the iterator to point to the previous cell and
    /// return a copy of the iterator before the decrement.
    /// </summary>
    cell_iterator operator--(int);

    /// <summary>
    /// Pre-increments the iterator to point to the previous cell and
    /// returns a reference to the iterator.
    /// </summary>
    cell_iterator &operator++();

    /// <summary>
    /// Post-increments the iterator to point to the previous cell and
    /// return a copy of the iterator before the decrement.
    /// </summary>
    cell_iterator operator++(int);

private:

    /// <summary>
    /// If true, cells that don't exist in the worksheet will be skipped during iteration.
    /// </summary>
    bool skip_null_ = false;

    /// <summary>
    /// If true, when on the last column, the cursor will continue to the next row
    /// (and vice versa when iterating in column-major order) until reaching the
    /// bottom right corner of the range.
    /// </summary>
    bool wrap_ = false;

    /// <summary>
    /// The order this iterator will move, by column or by row. Note that
    /// this has the opposite meaning as in a range_iterator because after
    /// getting a row-major range_iterator, the row-major cell_iterator will
    /// iterate over a column and vice versa.
    /// </summary>
    major_order order_ = major_order::column;

    /// <summary>
    /// The worksheet this iterator will return cells from.
    /// </summary>
    worksheet ws_;

    /// <summary>
    /// The current cell the iterator points to
    /// </summary>
    cell_reference cursor_;

    /// <summary>
    /// The range of cells this iterator is restricted to
    /// </summary>
    range_reference bounds_;
};

/// <summary>
/// A cell iterator iterates over a 1D range by row or by column.
/// </summary>
class XLNT_API const_cell_iterator
{
public:
    /// <summary>
    /// iterator tags required for use with standard algorithms and adapters
    /// </summary>
    using iterator_category = std::bidirectional_iterator_tag;
    using value_type = const cell;
    using difference_type = std::ptrdiff_t;
    using pointer = const cell *;
    using reference = const cell; // intentionally value

    /// <summary>
    /// Default constructs a cell_iterator
    /// </summary>
    const_cell_iterator() = default;

    /// <summary>
    /// Constructs a cell_iterator from a worksheet, range, and iteration settings.
    /// </summary>
    const_cell_iterator(worksheet ws, const cell_reference &start_cell,
        const range_reference &limits, major_order order, bool skip_null, bool wrap);

    /// <summary>
    /// Constructs a const_cell_iterator as a copy of an existing cell_iterator.
    /// </summary>
    const_cell_iterator(const const_cell_iterator &) = default;

    /// <summary>
    /// Assigns this iterator to match the data in rhs.
    /// </summary>
    const_cell_iterator &operator=(const const_cell_iterator &) = default;

    /// <summary>
    /// Constructs a const_cell_iterator by moving from a const_cell_iterator temporary
    /// </summary>
    const_cell_iterator(const_cell_iterator &&) = default;

    /// <summary>
    /// Assigns this iterator to from a const_cell_iterator temporary
    /// </summary>
    const_cell_iterator &operator=(const_cell_iterator &&) = default;

    /// <summary>
    /// destructor for const_cell_iterator
    /// </summary>
    ~const_cell_iterator() = default;

    /// <summary>
    /// Dereferences this iterator to return the cell it points to.
    /// </summary>
    const reference operator*() const;

    /// <summary>
    /// Returns true if this iterator is equivalent to other.
    /// </summary>
    bool operator==(const const_cell_iterator &other) const;

    /// <summary>
    /// Returns true if this iterator isn't equivalent to other.
    /// </summary>
    bool operator!=(const const_cell_iterator &other) const;

    /// <summary>
    /// Pre-decrements the iterator to point to the previous cell and
    /// returns a reference to the iterator.
    /// </summary>
    const_cell_iterator &operator--();

    /// <summary>
    /// Post-decrements the iterator to point to the previous cell and
    /// return a copy of the iterator before the decrement.
    /// </summary>
    const_cell_iterator operator--(int);

    /// <summary>
    /// Pre-increments the iterator to point to the previous cell and
    /// returns a reference to the iterator.
    /// </summary>
    const_cell_iterator &operator++();

    /// <summary>
    /// Post-increments the iterator to point to the previous cell and
    /// return a copy of the iterator before the decrement.
    /// </summary>
    const_cell_iterator operator++(int);

private:
    /// <summary>
    /// If true, cells that don't exist in the worksheet will be skipped during iteration.
    /// </summary>
    bool skip_null_ = false;

    /// <summary>
    /// If true, when on the last column, the cursor will continue to the next row
    /// (and vice versa when iterating in column-major order) until reaching the
    /// bottom right corner of the range.
    /// </summary>
    bool wrap_ = false;

    /// <summary>
    /// The order this iterator will move, by column or by row. Note that
    /// this has the opposite meaning as in a range_iterator because after
    /// getting a row-major range_iterator, the row-major cell_iterator will
    /// iterate over a column and vice versa.
    /// </summary>
    major_order order_ = major_order::column;

    /// <summary>
    /// The worksheet this iterator will return cells from.
    /// </summary>
    worksheet ws_;

    /// <summary>
    /// The current cell the iterator points to
    /// </summary>
    cell_reference cursor_;

    /// <summary>
    /// The range of cells this iterator is restricted to
    /// </summary>
    range_reference bounds_;
};

} // namespace xlnt
