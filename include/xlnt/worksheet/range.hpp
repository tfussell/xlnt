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

#include <iterator>
#include <memory>
#include <string>
#include <vector>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/worksheet/cell_vector.hpp>
#include <xlnt/worksheet/major_order.hpp>
#include <xlnt/worksheet/range_iterator_2d.hpp>
#include <xlnt/worksheet/range_reference.hpp>
#include <xlnt/worksheet/worksheet.hpp>

namespace xlnt {

/// <summary>
/// A range is a 2D collection of cells with defined extens that can be iterated upon.
/// </summary>
class XLNT_CLASS range
{
  public:
    using iterator = range_iterator_2d;
    using const_iterator = const_range_iterator_2d;
    
    range(worksheet ws, const range_reference &reference, major_order order = major_order::row, bool skip_null = false);

    ~range();

    cell_vector operator[](std::size_t vector_index);

    const cell_vector operator[](std::size_t vector_index) const;

    bool operator==(const range &comparand) const;

    bool operator!=(const range &comparand) const;

    cell_vector get_vector(std::size_t vector_index);

    const cell_vector get_vector(std::size_t vector_index) const;

    cell get_cell(const cell_reference &ref);

    const cell get_cell(const cell_reference &ref) const;

    range_reference get_reference() const;

    std::size_t length() const;

    bool contains(const cell_reference &ref);

    iterator begin();
    iterator end();

    const_iterator begin() const;
    const_iterator end() const;

    const_iterator cbegin() const;
    const_iterator cend() const;

  private:
    worksheet ws_;
    range_reference ref_;
    major_order order_;
    bool skip_null_;
};

} // namespace xlnt
