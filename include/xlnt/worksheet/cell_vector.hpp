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

#pragma once

#include <iterator>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/cell/cell_reference.hpp>
#include <xlnt/worksheet/cell_iterator.hpp>
#include <xlnt/worksheet/const_cell_iterator.hpp>
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
    ///
    /// </summary>
    using iterator = cell_iterator;

    /// <summary>
    ///
    /// </summary>
    using const_iterator = const_cell_iterator;

    /// <summary>
    ///
    /// </summary>
    using reverse_iterator = std::reverse_iterator<iterator>;

    /// <summary>
    ///
    /// </summary>
    using const_reverse_iterator = std::reverse_iterator<const_iterator>;

    /// <summary>
    ///
    /// </summary>
    cell_vector(worksheet ws, const range_reference &ref, major_order order = major_order::row);

    /// <summary>
    ///
    /// </summary>
    cell front();

    /// <summary>
    ///
    /// </summary>
    const cell front() const;

    /// <summary>
    ///
    /// </summary>
    cell back();

    /// <summary>
    ///
    /// </summary>
    const cell back() const;

    /// <summary>
    ///
    /// </summary>
    cell operator[](std::size_t column_index);

    /// <summary>
    ///
    /// </summary>
    const cell operator[](std::size_t column_index) const;

    /// <summary>
    ///
    /// </summary>
    std::size_t length() const;

    /// <summary>
    ///
    /// </summary>
    iterator begin();

    /// <summary>
    ///
    /// </summary>
    iterator end();

    /// <summary>
    ///
    /// </summary>
    const_iterator begin() const;

    /// <summary>
    ///
    /// </summary>
    const_iterator cbegin() const;

    /// <summary>
    ///
    /// </summary>
    const_iterator end() const;

    /// <summary>
    ///
    /// </summary>
    const_iterator cend() const;

    /// <summary>
    ///
    /// </summary>
    reverse_iterator rbegin();

    /// <summary>
    ///
    /// </summary>
    reverse_iterator rend();

    /// <summary>
    ///
    /// </summary>
    const_reverse_iterator rbegin() const;

    /// <summary>
    ///
    /// </summary>
    const_reverse_iterator rend() const;

    /// <summary>
    ///
    /// </summary>
    const_reverse_iterator crbegin() const;

    /// <summary>
    ///
    /// </summary>
    const_reverse_iterator crend() const;

private:
    /// <summary>
    ///
    /// </summary>
    worksheet ws_;

    /// <summary>
    ///
    /// </summary>
    range_reference ref_;

    /// <summary>
    ///
    /// </summary>
    major_order order_;
};

} // namespace xlnt
