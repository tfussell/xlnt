// Copyright (c) 2014-2021 Thomas Fussell
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

#include <string>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/cell/cell_reference.hpp>
#include <xlnt/utils/optional.hpp>

namespace xlnt {

struct XLNT_API sheet_pr
{
    /// <summary>
    /// is horizontally synced to the anchor point
    /// </summary>
    optional<bool> sync_horizontal;

    /// <summary>
    /// is vertically synced to the anchor point
    /// </summary>
    optional<bool> sync_vertical;

    /// <summary>
    /// Anchor point for worksheet's window
    /// </summary>
    optional<cell_reference> sync_ref;

    /// <summary>
    /// Lotus compatibility option
    /// </summary>
    optional<bool> transition_evaluation;

    /// <summary>
    /// Lotus compatibility option
    /// </summary>
    optional<bool> transition_entry;

    /// <summary>
    /// worksheet is published
    /// </summary>
    optional<bool> published;

    /// <summary>
    /// stable name of the sheet
    /// </summary>
    optional<std::string> code_name;

    /// <summary>
    /// worksheet has one or more autofilters or advanced filters on
    /// </summary>
    optional<bool> filter_mode;

    /// <summary>
    /// whether the conditional formatting calculations shall be evaluated
    /// </summary>
    optional<bool> enable_format_condition_calculation;
};

inline bool operator==(const sheet_pr &lhs, const sheet_pr &rhs)
{
    return lhs.sync_horizontal == rhs.sync_horizontal
        && lhs.sync_vertical == rhs.sync_vertical
        && lhs.sync_ref == rhs.sync_ref
        && lhs.transition_evaluation == rhs.transition_evaluation
        && lhs.transition_entry == rhs.transition_entry
        && lhs.published == rhs.published
        && lhs.code_name == rhs.code_name
        && lhs.filter_mode == rhs.filter_mode
        && lhs.enable_format_condition_calculation == rhs.enable_format_condition_calculation;
}
} // namespace xlnt
