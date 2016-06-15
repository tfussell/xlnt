// Copyright (c) 2014-2016 Thomas Fussell
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

class XLNT_CLASS const_worksheet_iterator : public std::iterator<std::bidirectional_iterator_tag, const worksheet, std::ptrdiff_t, const worksheet*, const worksheet>
{
public:
    const_worksheet_iterator(const workbook &wb, std::size_t index);
    const_worksheet_iterator(const const_worksheet_iterator &);
    const_worksheet_iterator &operator=(const const_worksheet_iterator &);
    const worksheet operator*();
    bool operator==(const const_worksheet_iterator &comparand) const;
    bool operator!=(const const_worksheet_iterator &comparand) const
    {
        return !(*this == comparand);
    }
    const_worksheet_iterator operator++(int);
    const_worksheet_iterator &operator++();

private:
    const workbook &wb_;
    std::size_t index_;
};

} // namespace xlnt
