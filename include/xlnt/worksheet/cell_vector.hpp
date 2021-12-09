// Copyright (c) 2014-2021 Thomas Fussell
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
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE
//
// @license: http://www.opensource.org/licenses/mit-license.php
// @author: see AUTHORS file

#pragma once

#include <iterator>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/cell/cell_reference.hpp>
#include <xlnt/worksheet/cell_iterator.hpp>
#include <xlnt/worksheet/major_order.hpp>
#include <xlnt/worksheet/range_reference.hpp>
#include <xlnt/worksheet/worksheet.hpp>

namespace xlnt {

class cell;
class cell_iterator;
class const_cell_iterator;
class range_reference;

/// <summary>
/// A cell vector is a linear (1D) range of cells, either vertical or horizontal
/// depending on the major order specified in the constructor.
/// </summary>
class XLNT_API cell_vector
{
public:
    /// <summary>
    /// Iterate over cells in a cell_vector with an iterator of this type.
    /// </summary>
    using iterator = cell_iterator;

    /// <summary>
    /// Iterate over const cells in a const cell_vector with an iterator of this type.
    /// </summary>
    using const_iterator = const_cell_iterator;

    /// <summary>
    /// Iterate over cells in a cell_vector in reverse oreder with an iterator
    /// of this type.
    /// </summary>
    using reverse_iterator = std::reverse_iterator<iterator>;

    /// <summary>
    /// Iterate over const cells in a const cell_vector in reverse order with
    /// an iterator of this type.
    /// </summary>
    using const_reverse_iterator = std::reverse_iterator<const_iterator>;

    /// <summary>
    /// Constructs a cell vector pointing to a given range in a given worksheet.
    /// order determines whether this vector is a row or a column. If skip_null is
    /// true, iterating over this vector will skip empty cells.
    /// </summary>
    cell_vector(worksheet ws, const cell_reference &cursor,
        const range_reference &ref, major_order order, bool skip_null, bool wrap);

    /// <summary>
    /// Returns true if every cell in this vector is null (i.e. the cell doesn't exist in the worksheet).
    /// </summary>
    bool empty() const;

    /// <summary>
    /// Returns the first cell in this vector.
    /// </summary>
    cell front();

    /// <summary>
    /// Returns the first cell in this vector.
    /// </summary>
    const cell front() const;

    /// <summary>
    /// Returns the last cell in this vector.
    /// </summary>
    cell back();

    /// <summary>
    /// Returns the last cell in this vector.
    /// </summary>
    const cell back() const;

    /// <summary>
    /// Returns the distance between the first and last cells in this vector.
    /// </summary>
    std::size_t length() const;

    /// <summary>
    /// Returns an iterator to the first cell in this vector.
    /// </summary>
    iterator begin();

    /// <summary>
    /// Returns an iterator to a cell one-past-the-end of this vector.
    /// </summary>
    iterator end();

    /// <summary>
    /// Returns an iterator to the first cell in this vector.
    /// </summary>
    const_iterator begin() const;

    /// <summary>
    /// Returns an iterator to the first cell in this vector.
    /// </summary>
    const_iterator cbegin() const;

    /// <summary>
    /// Returns an iterator to a cell one-past-the-end of this vector.
    /// </summary>
    const_iterator end() const;

    /// <summary>
    /// Returns an iterator to a cell one-past-the-end of this vector.
    /// </summary>
    const_iterator cend() const;

    /// <summary>
    /// Returns a reverse iterator to the first cell of this reversed vector.
    /// </summary>
    reverse_iterator rbegin();

    /// <summary>
    /// Returns a reverse iterator to to a cell one-past-the-end of this reversed vector.
    /// </summary>
    reverse_iterator rend();

    /// <summary>
    /// Returns a reverse iterator to the first cell of this reversed vector.
    /// </summary>
    const_reverse_iterator rbegin() const;

    /// <summary>
    /// Returns a reverse iterator to to a cell one-past-the-end of this reversed vector.
    /// </summary>
    const_reverse_iterator rend() const;

    /// <summary>
    /// Returns a reverse iterator to the first cell of this reversed vector.
    /// </summary>
    const_reverse_iterator crbegin() const;

    /// <summary>
    /// Returns a reverse iterator to to a cell one-past-the-end of this reversed vector.
    /// </summary>
    const_reverse_iterator crend() const;

    /// <summary>
    /// Returns the cell column_index distance away from the first cell in this vector.
    /// </summary>
    cell operator[](std::size_t column_index);

    /// <summary>
    /// Returns the cell column_index distance away from the first cell in this vector.
    /// </summary>
    const cell operator[](std::size_t column_index) const;

private:
    /// <summary>
    /// The worksheet this vector points to cells from
    /// </summary>
    worksheet ws_;

    /// <summary>
    /// The reference of the first cell in this vector
    /// </summary>
    cell_reference cursor_;

    /// <summary>
    /// The range of cells this vector can point to
    /// </summary>
    range_reference bounds_;

    /// <summary>
    /// The direction that iteration over this vector will move. Note that
    /// this has the opposite meaning as in a range_iterator because after
    /// getting a row-major range_iterator, the row-major cell_iterator will
    /// iterate over a column and vice versa.
    /// </summary>
    major_order order_;

    /// <summary>
    /// If true, cells that don't exist in the worksheet will be skipped during iteration.
    /// </summary>
    bool skip_null_;

    /// <summary>
    /// If true, when on the last column, the cursor will continue to the next row
    /// (and vice versa when iterating in column-major order) until reaching the
    /// bottom right corner of the range.
    /// </summary>
    bool wrap_;
};

} // namespace xlnt
