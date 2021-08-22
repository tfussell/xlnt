// Copyright (c) 2014-2021 Thomas Fussell
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
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE
//
// @license: http://www.opensource.org/licenses/mit-license.php
// @author: see AUTHORS file
#pragma once

#include <cstddef>
#include <iterator>

#include <xlnt/xlnt_config.hpp>

namespace xlnt {

class workbook;
class worksheet;

// Note: There are const and non-const implementations of this iterator
// because one needs to point at a const workbook and the other needs
// to point at a non-const workbook stored as a member variable, respectively.

/// <summary>
/// An iterator which is used to iterate over the worksheets in a workbook.
/// </summary>
class XLNT_API worksheet_iterator
{
public:
    /// <summary>
    /// iterator tags required for use with standard algorithms and adapters
    /// </summary>
    using iterator_category = std::bidirectional_iterator_tag;
    using value_type = worksheet;
    using difference_type = std::ptrdiff_t;
    using pointer = worksheet *;
    using reference = worksheet; // intentionally value

    /// <summary>
    /// Default Constructs a worksheet iterator
    /// </summary>
    worksheet_iterator() = default;

    /// <summary>
    /// Constructs a worksheet iterator from a workbook and sheet index.
    /// </summary>
    worksheet_iterator(workbook &wb, std::size_t index);

    /// <summary>
    /// Copy constructs a worksheet iterator from another iterator.
    /// </summary>
    worksheet_iterator(const worksheet_iterator &) = default;

    /// <summary>
    /// Copy assigns the iterator so that it points to the same worksheet in the same workbook.
    /// </summary>
    worksheet_iterator &operator=(const worksheet_iterator &) = default;

    /// <summary>
    /// Move constructs a worksheet iterator from a temporary iterator.
    /// </summary>
    worksheet_iterator(worksheet_iterator &&) = default;

    /// <summary>
    /// Move assign the iterator from a temporary iterator
    /// </summary>
    worksheet_iterator &operator=(worksheet_iterator &&) = default;

    /// <summary>
    /// Default destructor
    /// </summary>
    ~worksheet_iterator() = default;

    /// <summary>
    /// Dereferences the iterator to return the worksheet it is pointing to.
    /// If the iterator points to one-past-the-end of the workbook, an invalid_parameter
    /// exception will be thrown.
    /// </summary>
    reference operator*();

    /// <summary>
    /// Dereferences the iterator to return the worksheet it is pointing to.
    /// If the iterator points to one-past-the-end of the workbook, an invalid_parameter
    /// exception will be thrown.
    /// </summary>
    const reference operator*() const;

    /// <summary>
    /// Returns true if this iterator points to the same worksheet as comparand.
    /// </summary>
    bool operator==(const worksheet_iterator &comparand) const;

    /// <summary>
    /// Returns true if this iterator doesn't point to the same worksheet as comparand.
    /// </summary>
    bool operator!=(const worksheet_iterator &comparand) const;

    /// <summary>
    /// Post-increment the iterator's internal workseet index. Returns a copy of the
    /// iterator as it was before being incremented.
    /// </summary>
    worksheet_iterator operator++(int);

    /// <summary>
    /// Pre-increment the iterator's internal workseet index. Returns a refernce
    /// to the same iterator.
    /// </summary>
    worksheet_iterator &operator++();

    /// <summary>
    /// Post-decrement the iterator's internal workseet index. Returns a copy of the
    /// iterator as it was before being incremented.
    /// </summary>
    worksheet_iterator operator--(int);

    /// <summary>
    /// Pre-decrement the iterator's internal workseet index. Returns a refernce
    /// to the same iterator.
    /// </summary>
    worksheet_iterator &operator--();

private:
    /// <summary>
    /// The target workbook of this iterator.
    /// </summary>
    workbook *wb_ = nullptr;

    /// <summary>
    /// The index of the worksheet in wb_ this iterator is currently pointing to.
    /// </summary>
    std::size_t index_ = 0;
};

/// <summary>
/// An iterator which is used to iterate over the worksheets in a const workbook.
/// </summary>
class XLNT_API const_worksheet_iterator
{
public:
    /// <summary>
    /// iterator tags required for use with standard algorithms and adapters
    /// </summary>
    using iterator_category = std::bidirectional_iterator_tag;
    using value_type = const worksheet;
    using difference_type = std::ptrdiff_t;
    using pointer = const worksheet *;
    using reference = const worksheet; // intentionally value

    /// <summary>
    /// Default Constructs a worksheet iterator
    /// </summary>
    const_worksheet_iterator() = default;

    /// <summary>
    /// Constructs a worksheet iterator from a workbook and sheet index.
    /// </summary>
    const_worksheet_iterator(const workbook &wb, std::size_t index);

    /// <summary>
    /// Copy constructs a worksheet iterator from another iterator.
    /// </summary>
    const_worksheet_iterator(const const_worksheet_iterator &) = default;

    /// <summary>
    /// Copy assigns the iterator so that it points to the same worksheet in the same workbook.
    /// </summary>
    const_worksheet_iterator &operator=(const const_worksheet_iterator &) = default;

    /// <summary>
    /// Move constructs a worksheet iterator from a temporary iterator.
    /// </summary>
    const_worksheet_iterator(const_worksheet_iterator &&) = default;

    /// <summary>
    /// Move assigns the iterator from a temporary iterator
    /// </summary>
    const_worksheet_iterator &operator=(const_worksheet_iterator &&) = default;

    /// <summary>
    /// Default destructor
    /// </summary>
    ~const_worksheet_iterator() = default;

    /// <summary>
    /// Dereferences the iterator to return the worksheet it is pointing to.
    /// If the iterator points to one-past-the-end of the workbook, an invalid_parameter
    /// exception will be thrown.
    /// </summary>
    const reference operator*() const;

    /// <summary>
    /// Returns true if this iterator points to the same worksheet as comparand.
    /// </summary>
    bool operator==(const const_worksheet_iterator &comparand) const;

    /// <summary>
    /// Returns true if this iterator doesn't point to the same worksheet as comparand.
    /// </summary>
    bool operator!=(const const_worksheet_iterator &comparand) const;

    /// <summary>
    /// Post-increment the iterator's internal workseet index. Returns a copy of the
    /// iterator as it was before being incremented.
    /// </summary>
    const_worksheet_iterator operator++(int);

    /// <summary>
    /// Pre-increment the iterator's internal workseet index. Returns a refernce
    /// to the same iterator.
    /// </summary>
    const_worksheet_iterator &operator++();

    /// <summary>
    /// Post-decrement the iterator's internal workseet index. Returns a copy of the
    /// iterator as it was before being incremented.
    /// </summary>
    const_worksheet_iterator operator--(int);

    /// <summary>
    /// Pre-decrement the iterator's internal workseet index. Returns a refernce
    /// to the same iterator.
    /// </summary>
    const_worksheet_iterator &operator--();

private:
    /// <summary>
    /// The target workbook of this iterator.
    /// </summary>
    const workbook *wb_ = nullptr;

    /// <summary>
    /// The index of the worksheet in wb_ this iterator is currently pointing to.
    /// </summary>
    std::size_t index_ = 0;
};

} // namespace xlnt
