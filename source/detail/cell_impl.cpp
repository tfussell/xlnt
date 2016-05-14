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
#include <xlnt/worksheet/worksheet.hpp>

#include "cell_impl.hpp"
#include "comment_impl.hpp"

namespace xlnt {
namespace detail {

cell_impl::cell_impl() : cell_impl(column_t(1), 1)
{
}

cell_impl::cell_impl(column_t column, row_t row) : cell_impl(nullptr, column, row)
{
}

cell_impl::cell_impl(worksheet_impl *parent, column_t column, row_t row)
    : type_(cell::type::null),
      parent_(parent),
      column_(column),
      row_(row),
      value_numeric_(0),
      has_hyperlink_(false),
      is_merged_(false),
      has_format_(false),
      format_id_(0),
      comment_(nullptr)
{
}

cell_impl::cell_impl(const cell_impl &rhs)
{
    *this = rhs;
}

cell_impl &cell_impl::operator=(const cell_impl &rhs)
{
    parent_ = rhs.parent_;
    value_numeric_ = rhs.value_numeric_;
    value_text_ = rhs.value_text_;
    hyperlink_ = rhs.hyperlink_;
    formula_ = rhs.formula_;
    column_ = rhs.column_;
    row_ = rhs.row_;
    is_merged_ = rhs.is_merged_;
    has_hyperlink_ = rhs.has_hyperlink_;
    type_ = rhs.type_;
    format_id_ = rhs.format_id_;
    has_format_ = rhs.has_format_;

    if (rhs.comment_ != nullptr)
    {
        comment_.reset(new comment_impl(*rhs.comment_));
    }

    return *this;
}

cell cell_impl::self()
{
    return xlnt::cell(this);
}

} // namespace detail
} // namespace xlnt
