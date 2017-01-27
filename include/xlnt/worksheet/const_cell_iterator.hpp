// Copyright (c) 2014-2017 Thomas Fussell
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
#include <xlnt/worksheet/worksheet.hpp>

namespace xlnt {

enum class major_order;

class cell;
class cell_reference;
class range_reference;

/// <summary>
/// Alias the parent class of this iterator to increase clarity.
/// </summary>
using cc_iter_type = std::iterator<std::bidirectional_iterator_tag,
    const cell, std::ptrdiff_t, const cell *, const cell>;

/// <summary>
///
/// </summary>
class XLNT_API const_cell_iterator : public cc_iter_type
{
public:
    /// <summary>
    ///
    /// </summary>
    const_cell_iterator(worksheet ws, const cell_reference &start_cell);

    /// <summary>
    ///
    /// </summary>
    const_cell_iterator(worksheet ws, const cell_reference &start_cell, major_order order);

    /// <summary>
    ///
    /// </summary>
    const_cell_iterator(const const_cell_iterator &other);

    /// <summary>
    ///
    /// </summary>
    const_cell_iterator &operator=(const const_cell_iterator &) = default;

    /// <summary>
    ///
    /// </summary>
    const cell operator*() const;

    /// <summary>
    ///
    /// </summary>
    bool operator==(const const_cell_iterator &other) const;

    /// <summary>
    ///
    /// </summary>
    bool operator!=(const const_cell_iterator &other) const;

    /// <summary>
    ///
    /// </summary>
    const_cell_iterator &operator--();

    /// <summary>
    ///
    /// </summary>
    const_cell_iterator operator--(int);

    /// <summary>
    ///
    /// </summary>
    const_cell_iterator &operator++();

    /// <summary>
    ///
    /// </summary>
    const_cell_iterator operator++(int);

private:
    /// <summary>
    ///
    /// </summary>
    worksheet ws_;

    /// <summary>
    ///
    /// </summary>
    cell_reference current_cell_;

    /// <summary>
    ///
    /// </summary>
    range_reference range_;

    /// <summary>
    ///
    /// </summary>
    major_order order_;
};

} // namespace xlnt
