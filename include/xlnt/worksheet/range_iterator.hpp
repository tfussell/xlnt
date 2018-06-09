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

#pragma once

#include <cstddef> // std::ptrdiff_t

#include <xlnt/xlnt_config.hpp>
#include <xlnt/cell/cell_reference.hpp>
#include <xlnt/worksheet/major_order.hpp>
#include <xlnt/worksheet/range_reference.hpp>
#include <xlnt/worksheet/worksheet.hpp>

namespace xlnt {

class cell_vector;

/// <summary>
/// An iterator used by worksheet and range for traversing
/// a 2D grid of cells by row/column then across that row/column.
/// </summary>
class XLNT_API range_iterator
{
public:
    /// <summary>
    /// iterator tags required for use with standard algorithms and adapters
    /// </summary>
    using iterator_category = std::bidirectional_iterator_tag;
    using value_type = cell_vector;
    using difference_type = std::ptrdiff_t;
    using pointer = cell_vector *;
    using reference = cell_vector; // intentionally value

    /// <summary>
    /// Constructs a range iterator on a worksheet, cell pointing to the current
    /// row or column, range bounds, an order, and whether or not to skip null column/rows.
    /// </summary>
    range_iterator(worksheet &ws, const cell_reference &cursor,
        const range_reference &bounds, major_order order, bool skip_null);

    /// <summary>
    /// Default copy constructor.
    /// </summary>
    range_iterator(const range_iterator &) = default;

    /// <summary>
    /// Default assignment operator.
    /// </summary>
    range_iterator &operator=(const range_iterator &) = default;

    /// <summary>
    /// Default move constructor.
    /// </summary>
    range_iterator(range_iterator &&) = default;

    /// <summary>
    /// Default move assignment operator.
    /// </summary>
    range_iterator &operator=(range_iterator &&) = default;

    /// <summary>
    /// Default destructor
    /// </summary>
    ~range_iterator() = default;

    /// <summary>
    /// Dereference the iterator to return a column or row.
    /// </summary>
    reference operator*();

    /// <summary>
    /// Dereference the iterator to return a column or row.
    /// </summary>
    const reference operator*() const;

    /// <summary>
    /// Returns true if this iterator is equivalent to other.
    /// </summary>
    bool operator==(const range_iterator &other) const;

    /// <summary>
    /// Returns true if this iterator is not equivalent to other.
    /// </summary>
    bool operator!=(const range_iterator &other) const;

    /// <summary>
    /// Pre-decrement the iterator to point to the previous row/column.
    /// </summary>
    range_iterator &operator--();

    /// <summary>
    /// Post-decrement the iterator to point to the previous row/column.
    /// </summary>
    range_iterator operator--(int);

    /// <summary>
    /// Pre-increment the iterator to point to the next row/column.
    /// </summary>
    range_iterator &operator++();

    /// <summary>
    /// Post-increment the iterator to point to the next row/column.
    /// </summary>
    range_iterator operator++(int);

private:
    /// <summary>
    /// The worksheet
    /// </summary>
    worksheet ws_;

    /// <summary>
    /// The first cell in the current row/column
    /// </summary>
    cell_reference cursor_;

    /// <summary>
    /// The bounds of the range
    /// </summary>
    range_reference bounds_;

    /// <summary>
    /// Whether rows or columns should be iterated over first
    /// </summary>
    major_order order_;

    /// <summary>
    /// If true, empty rows and cells will be skipped when iterating with this iterator
    /// </summary>
    bool skip_null_;
};

/// <summary>
/// A const version of range_iterator which does not allow modification
/// to the dereferenced cell_vector.
/// </summary>
class XLNT_API const_range_iterator
{
public:
    /// <summary>
    /// this iterator meets the interface requirements of bidirection_iterator
    /// </summary>
    using iterator_category = std::bidirectional_iterator_tag;
    using value_type = const cell_vector;
    using difference_type = std::ptrdiff_t;
    using pointer = const cell_vector *;
    using reference = const cell_vector; // intentionally value

    /// <summary>
    /// Constructs a range iterator on a worksheet, cell pointing to the current
    /// row or column, range bounds, an order, and whether or not to skip null column/rows.
    /// </summary>
    const_range_iterator(const worksheet &ws, const cell_reference &cursor,
        const range_reference &bounds, major_order order, bool skip_null);

    /// <summary>
    /// Default copy constructor.
    /// </summary>
    const_range_iterator(const const_range_iterator &) = default;

    /// <summary>
    /// Default assignment operator.
    /// </summary>
    const_range_iterator &operator=(const const_range_iterator &) = default;

    /// <summary>
    /// Default move constructor.
    /// </summary>
    const_range_iterator(const_range_iterator &&) = default;

    /// <summary>
    /// Default move assignment operator.
    /// </summary>
    const_range_iterator &operator=(const_range_iterator &&) = default;

    /// <summary>
    /// Default destructor
    /// </summary>
    ~const_range_iterator() = default;

    /// <summary>
    /// Dereferennce the iterator to return the current column/row.
    /// </summary>
    const reference operator*() const;

    /// <summary>
    /// Returns true if this iterator is equivalent to other.
    /// </summary>
    bool operator==(const const_range_iterator &other) const;

    /// <summary>
    /// Returns true if this iterator is not equivalent to other.
    /// </summary>
    bool operator!=(const const_range_iterator &other) const;

    /// <summary>
    /// Pre-decrement the iterator to point to the next row/column.
    /// </summary>
    const_range_iterator &operator--();

    /// <summary>
    /// Post-decrement the iterator to point to the next row/column.
    /// </summary>
    const_range_iterator operator--(int);

    /// <summary>
    /// Pre-increment the iterator to point to the next row/column.
    /// </summary>
    const_range_iterator &operator++();

    /// <summary>
    /// Post-increment the iterator to point to the next row/column.
    /// </summary>
    const_range_iterator operator++(int);

private:
    /// <summary>
    /// The implementation of the worksheet this iterator points to
    /// </summary>
    detail::worksheet_impl *ws_;

    /// <summary>
    /// The first cell in the current row or column this iterator points to
    /// </summary>
    cell_reference cursor_;

    /// <summary>
    /// The range this iterator starts and ends in
    /// </summary>
    range_reference bounds_;

    /// <summary>
    /// Determines whether iteration should move through rows or columns first
    /// </summary>
    major_order order_;

    /// <summary>
    /// If true, empty rows and cells will be skipped when iterating with this iterator
    /// </summary>
    bool skip_null_;
};

} // namespace xlnt
