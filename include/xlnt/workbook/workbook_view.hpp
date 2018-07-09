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

#include <xlnt/xlnt_config.hpp>
#include <xlnt/utils/optional.hpp>

namespace xlnt {

/// <summary>
/// A workbook can be opened in multiple windows with different views.
/// This class represents a particular view used by one window.
/// </summary>
class XLNT_API workbook_view
{
public:
    /// <summary>
    /// If true, dates will be grouped when presenting the user with filtering options.
    /// </summary>
    bool auto_filter_date_grouping = true;

    /// <summary>
    /// If true, the view will be minimized.
    /// </summary>
    bool minimized = false;

    /// <summary>
    /// If true, the horizontal scroll bar will be displayed.
    /// </summary>
    bool show_horizontal_scroll = true;

    /// <summary>
    /// If true, the sheet tabs will be displayed.
    /// </summary>
    bool show_sheet_tabs = true;

    /// <summary>
    /// If true, the vertical scroll bar will be displayed.
    /// </summary>
    bool show_vertical_scroll = true;

    /// <summary>
    /// If true, the workbook window will be visible.
    /// </summary>
    bool visible = true;

    /// <summary>
    /// The optional index to the active sheet in this view.
    /// </summary>
    optional<std::size_t> active_tab;

    /// <summary>
    /// The optional index to the first sheet in this view.
    /// </summary>
    optional<std::size_t> first_sheet;

    /// <summary>
    /// The optional ratio between the tabs bar and the horizontal scroll bar.
    /// </summary>
    optional<std::size_t> tab_ratio;

    /// <summary>
    /// The width of the workbook window in twips.
    /// </summary>
    optional<std::size_t> window_width;

    /// <summary>
    /// The height of the workbook window in twips.
    /// </summary>
    optional<std::size_t> window_height;

    /// <summary>
    /// The distance of the workbook window from the left side of the screen in twips.
    /// </summary>
    optional<int> x_window;

    /// <summary>
    /// The distance of the workbook window from the top of the screen in twips.
    /// </summary>
    optional<int> y_window;
};

inline bool operator==(const workbook_view &lhs, const workbook_view &rhs)
{
    return lhs.active_tab == rhs.active_tab
        && lhs.auto_filter_date_grouping == rhs.auto_filter_date_grouping
        && lhs.first_sheet == rhs.first_sheet
        && lhs.minimized == rhs.minimized
        && lhs.show_horizontal_scroll == rhs.show_horizontal_scroll
        && lhs.show_sheet_tabs == rhs.show_sheet_tabs
        && lhs.show_vertical_scroll == rhs.show_vertical_scroll
        && lhs.tab_ratio == rhs.tab_ratio
        && lhs.visible == rhs.visible;
}

} // namespace xlnt
