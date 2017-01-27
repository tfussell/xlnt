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
using ws_iter_type = std::iterator<std::bidirectional_iterator_tag,
    worksheet, std::ptrdiff_t, worksheet *, worksheet>;

/// <summary>
///
/// </summary>
class XLNT_API worksheet_iterator : public ws_iter_type
{
public:
    /// <summary>
    ///
    /// </summary>
    worksheet_iterator(workbook &wb, std::size_t index);

    /// <summary>
    ///
    /// </summary>
    worksheet_iterator(const worksheet_iterator &);

    /// <summary>
    ///
    /// </summary>
    worksheet_iterator &operator=(const worksheet_iterator &);

    /// <summary>
    ///
    /// </summary>
    worksheet operator*();

    /// <summary>
    ///
    /// </summary>
    bool operator==(const worksheet_iterator &comparand) const;

    /// <summary>
    ///
    /// </summary>
    bool operator!=(const worksheet_iterator &comparand) const
    {
        return !(*this == comparand);
    }

    /// <summary>
    ///
    /// </summary>
    worksheet_iterator operator++(int);

    /// <summary>
    ///
    /// </summary>
    worksheet_iterator &operator++();

private:
    /// <summary>
    ///
    /// </summary>
    workbook &wb_;

    /// <summary>
    ///
    /// </summary>
    std::size_t index_;
};

} // namespace xlnt
