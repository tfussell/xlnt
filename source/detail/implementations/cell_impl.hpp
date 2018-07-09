// Copyright (c) 2014-2018 Thomas Fussell
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
#include <xlnt/cell/index_types.hpp>
#include <xlnt/cell/rich_text.hpp>
#include <xlnt/packaging/relationship.hpp>
#include <xlnt/utils/optional.hpp>
#include <detail/implementations/format_impl.hpp>
#include <detail/implementations/hyperlink_impl.hpp>

namespace xlnt {
namespace detail {

struct worksheet_impl;

struct cell_impl
{
    cell_impl();
    cell_impl(const cell_impl &other) = default;
    cell_impl(cell_impl &&other) = default;
    cell_impl &operator=(const cell_impl &other) = default;
    cell_impl &operator=(cell_impl &&other) = default;
    
    cell_type type_;

    worksheet_impl *parent_;

    column_t column_;
    row_t row_;

    bool is_merged_;

    rich_text value_text_;
    double value_numeric_;

    optional<std::string> formula_;
    optional<hyperlink_impl> hyperlink_;
    optional<format_impl *> format_;
    optional<comment *> comment_;
};

inline bool operator==(const cell_impl &lhs, const cell_impl &rhs)
{
    // not comparing parent
    return lhs.type_ == rhs.type_
        && lhs.column_ == rhs.column_
        && lhs.row_ == rhs.row_
        && lhs.is_merged_ == rhs.is_merged_
        && lhs.value_text_ == rhs.value_text_
        && lhs.value_numeric_ == rhs.value_numeric_
        && lhs.formula_ == rhs.formula_
        && lhs.hyperlink_ == rhs.hyperlink_
        && (lhs.format_.is_set() == rhs.format_.is_set() && (!lhs.format_.is_set() || *lhs.format_.get() == *rhs.format_.get()))
        && (lhs.comment_.is_set() == rhs.comment_.is_set() && (!lhs.comment_.is_set() || *lhs.comment_.get() == *rhs.comment_.get()));
}

} // namespace detail
} // namespace xlnt
