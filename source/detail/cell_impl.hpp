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

#include <cstdlib>

#include <xlnt/cell/cell.hpp>
#include <xlnt/cell/comment.hpp>
#include <xlnt/cell/index_types.hpp>
#include <xlnt/utils/exceptions.hpp>
#include <xlnt/utils/time.hpp>
#include <xlnt/packaging/relationship.hpp>
#include <xlnt/styles/number_format.hpp>

#include "comment_impl.hpp"

namespace xlnt {

class style;

namespace detail {

struct worksheet_impl;

struct cell_impl
{
    cell_impl();
    cell_impl(column_t column, row_t row);
    cell_impl(worksheet_impl *parent, column_t column, row_t row);
    cell_impl(const cell_impl &rhs);
    cell_impl &operator=(const cell_impl &rhs);

    cell self();

    void set_string(const std::string &s, bool guess_types);

    cell::type type_;

    worksheet_impl *parent_;

    column_t column_;
    row_t row_;

    std::string value_string_;
    long double value_numeric_;

    std::string formula_;

    bool has_hyperlink_;
    relationship hyperlink_;

    bool is_merged_;

    bool has_style_;
    std::size_t style_id_;

    std::unique_ptr<comment_impl> comment_;
};

} // namespace detail
} // namespace xlnt
