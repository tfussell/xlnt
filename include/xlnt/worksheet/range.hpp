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

#include <functional>
#include <iterator>
#include <memory>
#include <string>
#include <vector>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/styles/alignment.hpp>
#include <xlnt/styles/border.hpp>
#include <xlnt/styles/conditional_format.hpp>
#include <xlnt/styles/fill.hpp>
#include <xlnt/styles/font.hpp>
#include <xlnt/styles/number_format.hpp>
#include <xlnt/styles/protection.hpp>
#include <xlnt/worksheet/cell_vector.hpp>
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
    /// Alias for the iterator type
    /// </summary>
    using iterator = range_iterator;

    /// <summary>
    /// Alias for the const iterator type
    /// </summary>
    using const_iterator = const_range_iterator;

    /// <summary>
    /// Alias for the reverse iterator type
    /// </summary>
    using reverse_iterator = std::reverse_iterator<iterator>;

    /// <summary>
    /// Alias for the const reverse iterator type
    /// </summary>
    using const_reverse_iterator = std::reverse_iterator<const_iterator>;

    /// <summary>
    /// Constructs a range on the given worksheet.
    /// </summary>
    range(worksheet ws, const range_reference &reference,
        major_order order = major_order::row, bool skip_null = false);

    /// <summary>
    /// Desctructor
    /// </summary>
    ~range();

    /// <summary>
    /// Default copy constructor.
    /// </summary>
    range(const range &) = default;

    /// <summary>
    /// Erases all cell data from the worksheet for cells within this range.
    /// </summary>
    void clear_cells();

    /// <summary>
    /// Returns a vector pointing to the n-th row or column in this range (depending
    /// on major order).
    /// </summary>
    cell_vector vector(std::size_t n);

    /// <summary>
    /// Returns a vector pointing to the n-th row or column in this range (depending
    /// on major order).
    /// </summary>
    const cell_vector vector(std::size_t n) const;

    /// <summary>
    /// Returns a cell in the range relative to its top left cell.
    /// </summary>
    class cell cell(const cell_reference &ref);

    /// <summary>
    /// Returns a cell in the range relative to its top left cell.
    /// </summary>
    const class cell cell(const cell_reference &ref) const;

    /// <summary>
    /// The worksheet this range targets
    /// </summary>
    const worksheet &target_worksheet() const;

    /// <summary>
    /// Returns the reference defining the bounds of this range.
    /// </summary>
    range_reference reference() const;

    /// <summary>
    /// Returns the number of rows or columns in this range (depending on major order).
    /// </summary>
    std::size_t length() const;

    /// <summary>
    /// Returns true if the given cell exists in the parent worksheet of this range.
    /// </summary>
    bool contains(const cell_reference &ref);

    /// <summary>
    /// Sets the alignment of all cells in the range to new_alignment and returns the range.
    /// </summary>
    range alignment(const xlnt::alignment &new_alignment);

    /// <summary>
    /// Sets the border of all cells in the range to new_border and returns the range.
    /// </summary>
    range border(const xlnt::border &new_border);

    /// <summary>
    /// Sets the fill of all cells in the range to new_fill and returns the range.
    /// </summary>
    range fill(const xlnt::fill &new_fill);

    /// <summary>
    /// Sets the font of all cells in the range to new_font and returns the range.
    /// </summary>
    range font(const xlnt::font &new_font);

    /// <summary>
    /// Sets the number format of all cells in the range to new_number_format and
    /// returns the range.
    /// </summary>
    range number_format(const xlnt::number_format &new_number_format);

    /// <summary>
    /// Sets the protection of all cells in the range to new_protection and returns the range.
    /// </summary>
    range protection(const xlnt::protection &new_protection);

    /// <summary>
    /// Sets the named style applied to all cells in this range to a style named style_name.
    /// </summary>
    range style(const class style &new_style);

    /// <summary>
    /// Sets the named style applied to all cells in this range to a style named style_name.
    /// If this style has not been previously created in the workbook, a
    /// key_not_found exception will be thrown.
    /// </summary>
    range style(const std::string &style_name);

    /// <summary>
    ///
    /// </summary>
    xlnt::conditional_format conditional_format(const condition &when);

    /// <summary>
    /// Returns the first row or column in this range.
    /// </summary>
    cell_vector front();

    /// <summary>
    /// Returns the first row or column in this range.
    /// </summary>
    const cell_vector front() const;

    /// <summary>
    /// Returns the last row or column in this range.
    /// </summary>
    cell_vector back();

    /// <summary>
    /// Returns the last row or column in this range.
    /// </summary>
    const cell_vector back() const;

    /// <summary>
    /// Returns an iterator to the first row or column in this range.
    /// </summary>
    iterator begin();

    /// <summary>
    /// Returns an iterator to one past the last row or column in this range.
    /// </summary>
    iterator end();

    /// <summary>
    /// Returns an iterator to the first row or column in this range.
    /// </summary>
    const_iterator begin() const;

    /// <summary>
    /// Returns an iterator to one past the last row or column in this range.
    /// </summary>
    const_iterator end() const;

    /// <summary>
    /// Returns an iterator to the first row or column in this range.
    /// </summary>
    const_iterator cbegin() const;

    /// <summary>
    /// Returns an iterator to one past the last row or column in this range.
    /// </summary>
    const_iterator cend() const;

    /// <summary>
    /// Returns a reverse iterator to the first row or column in this range.
    /// </summary>
    reverse_iterator rbegin();

    /// <summary>
    /// Returns a reverse iterator to one past the last row or column in this range.
    /// </summary>
    reverse_iterator rend();

    /// <summary>
    /// Returns a reverse iterator to the first row or column in this range.
    /// </summary>
    const_reverse_iterator rbegin() const;

    /// <summary>
    /// Returns a reverse iterator to one past the last row or column in this range.
    /// </summary>
    const_reverse_iterator rend() const;

    /// <summary>
    /// Returns a reverse iterator to the first row or column in this range.
    /// </summary>
    const_reverse_iterator crbegin() const;

    /// <summary>
    /// Returns a reverse iterator to one past the last row or column in this range.
    /// </summary>
    const_reverse_iterator crend() const;

    /// <summary>
    /// Applies function f to all cells in the range
    /// </summary>
    void apply(std::function<void(class cell)> f);

    /// <summary>
    /// Returns the n-th row or column in this range.
    /// </summary>
    cell_vector operator[](std::size_t n);

    /// <summary>
    /// Returns the n-th row or column in this range.
    /// </summary>
    const cell_vector operator[](std::size_t n) const;

    /// <summary>
    /// Returns true if this range is equivalent to comparand.
    /// </summary>
    bool operator==(const range &comparand) const;

    /// <summary>
    /// Returns true if this range is not equivalent to comparand.
    /// </summary>
    bool operator!=(const range &comparand) const;

private:
    /// <summary>
    /// The worksheet this range is within
    /// </summary>
    class worksheet ws_;

    /// <summary>
    /// The reference of this range
    /// </summary>
    range_reference ref_;

    /// <summary>
    /// Whether this range should be iterated by columns or rows first
    /// </summary>
    major_order order_;

    /// <summary>
    /// Whether null rows/columns and cells should be skipped during iteration
    /// </summary>
    bool skip_null_;
};

} // namespace xlnt
