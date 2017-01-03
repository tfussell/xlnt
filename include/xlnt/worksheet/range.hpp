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
#include <memory>
#include <string>
#include <vector>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/worksheet/cell_vector.hpp>
#include <xlnt/worksheet/const_range_iterator.hpp>
#include <xlnt/worksheet/major_order.hpp>
#include <xlnt/worksheet/range_iterator.hpp>
#include <xlnt/worksheet/range_reference.hpp>
#include <xlnt/worksheet/worksheet.hpp>

namespace xlnt {

class const_range_iterator;
class range_iterator;

/// <summary>
/// A range is a 2D collection of cells with defined extens that can be iterated upon.
/// </summary>
class XLNT_API range
{
public:
    /// <summary>
    ///
    /// </summary>
    using iterator = range_iterator;

    /// <summary>
    ///
    /// </summary>
    using const_iterator = const_range_iterator;

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
    range(worksheet ws, const range_reference &reference, major_order order = major_order::row, bool skip_null = false);

    /// <summary>
    ///
    /// </summary>
    ~range();

    /// <summary>
    ///
    /// </summary>
    range(const range &) = default;

    /// <summary>
    ///
    /// </summary>
    cell_vector operator[](std::size_t vector_index);

    /// <summary>
    ///
    /// </summary>
    const cell_vector operator[](std::size_t vector_index) const;

    /// <summary>
    ///
    /// </summary>
    bool operator==(const range &comparand) const;

    /// <summary>
    ///
    /// </summary>
    bool operator!=(const range &comparand) const;

    /// <summary>
    ///
    /// </summary>
    cell_vector vector(std::size_t vector_index);

    /// <summary>
    ///
    /// </summary>
    const cell_vector vector(std::size_t vector_index) const;

    /// <summary>
    ///
    /// </summary>
    class cell cell(const cell_reference &ref);

    /// <summary>
    ///
    /// </summary>
    const class cell cell(const cell_reference &ref) const;

    /// <summary>
    ///
    /// </summary>
    range_reference reference() const;

    /// <summary>
    ///
    /// </summary>
    std::size_t length() const;

    /// <summary>
    ///
    /// </summary>
    bool contains(const cell_reference &ref);

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
    const_iterator end() const;

    /// <summary>
    ///
    /// </summary>
    const_iterator cbegin() const;

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

    /// <summary>
    ///
    /// </summary>
    bool skip_null_;
};

} // namespace xlnt
