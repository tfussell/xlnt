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
#include <xlnt/workbook/const_worksheet_iterator.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/worksheet/worksheet.hpp>

namespace xlnt {

const_worksheet_iterator::const_worksheet_iterator(const workbook &wb, std::size_t index)
    : wb_(wb), index_(index)
{
}

const_worksheet_iterator::const_worksheet_iterator(const const_worksheet_iterator &rhs)
    : wb_(rhs.wb_), index_(rhs.index_)
{
}

const worksheet const_worksheet_iterator::operator*()
{
    return wb_.sheet_by_index(index_);
}

const_worksheet_iterator &const_worksheet_iterator::operator++()
{
    index_++;
    return *this;
}

const_worksheet_iterator const_worksheet_iterator::operator++(int)
{
    const_worksheet_iterator old(wb_, index_);
    ++*this;
    return old;
}

bool const_worksheet_iterator::operator==(const const_worksheet_iterator &comparand) const
{
    return index_ == comparand.index_ && wb_ == comparand.wb_;
}

} // namespace xlnt
