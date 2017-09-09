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

#include <cstddef>
#include <string>

#include <xlnt/cell/cell_type.hpp>
#include <xlnt/cell/comment.hpp>
#include <xlnt/cell/rich_text.hpp>
#include <xlnt/cell/index_types.hpp>
#include <xlnt/utils/optional.hpp>

namespace xlnt {
namespace detail {

struct format_impl;
struct worksheet_impl;

struct cell_impl
{
    cell_impl();

    cell_type type_;

    worksheet_impl *parent_;

    column_t column_;
    row_t row_;

    bool is_merged_;

    rich_text value_text_;
    double value_numeric_;

    optional<std::string> formula_;
    optional<std::string> hyperlink_;
    optional<format_impl *> format_;
    optional<comment *> comment_;
};

} // namespace detail
} // namespace xlnt
