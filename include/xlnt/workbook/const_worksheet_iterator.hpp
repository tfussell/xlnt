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

#include <cstddef>
#include <iterator>

#include <xlnt/xlnt_config.hpp>

namespace xlnt {

class workbook;
class worksheet;

/// <summary>
/// Alias the parent class of this iterator to increase clarity.
/// </summary>
using c_ws_iter_type = std::iterator<std::bidirectional_iterator_tag,
    const worksheet, std::ptrdiff_t, const worksheet *, const worksheet>;

/// <summary>
///
/// </summary>
class XLNT_API const_worksheet_iterator : public c_ws_iter_type
{
public:
    /// <summary>
    ///
    /// </summary>
    const_worksheet_iterator(const workbook &wb, std::size_t index);

    /// <summary>
    ///
    /// </summary>
    const_worksheet_iterator(const const_worksheet_iterator &);

    /// <summary>
    ///
    /// </summary>
    const_worksheet_iterator &operator=(const const_worksheet_iterator &);

    /// <summary>
    ///
    /// </summary>
    const worksheet operator*();

    /// <summary>
    ///
    /// </summary>
    bool operator==(const const_worksheet_iterator &comparand) const;

    /// <summary>
    ///
    /// </summary>
    bool operator!=(const const_worksheet_iterator &comparand) const
    {
        return !(*this == comparand);
    }

    /// <summary>
    ///
    /// </summary>
    const_worksheet_iterator operator++(int);

    /// <summary>
    ///
    /// </summary>
    const_worksheet_iterator &operator++();

private:
    /// <summary>
    ///
    /// </summary>
    const workbook &wb_;

    /// <summary>
    ///
    /// </summary>
    std::size_t index_;
};

} // namespace xlnt
