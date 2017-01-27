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

#include <xlnt/xlnt_config.hpp>
#include <xlnt/utils/optional.hpp>
#include <xlnt/worksheet/pane.hpp>
#include <xlnt/worksheet/selection.hpp>

namespace xlnt {

/// <summary>
///
/// </summary>
enum class sheet_view_type
{
    normal,
    page_break_preview,
    page_layout
};

/// <summary>
/// Describes a view of a worksheet.
/// Worksheets can have multiple views which show the data differently.
/// </summary>
class XLNT_API sheet_view
{
public:
    /// <summary>
    ///
    /// </summary>
    void id(std::size_t new_id)
    {
        id_ = new_id;
    }

    /// <summary>
    ///
    /// </summary>
    std::size_t id() const
    {
        return id_;
    }

    /// <summary>
    ///
    /// </summary>
    bool has_pane() const
    {
        return pane_.is_set();
    }

    /// <summary>
    ///
    /// </summary>
    struct pane &pane()
    {
        return pane_.get();
    }

    /// <summary>
    ///
    /// </summary>
    const struct pane &pane() const
    {
        return pane_.get();
    }

    /// <summary>
    ///
    /// </summary>
    void clear_pane()
    {
        pane_.clear();
    }

    /// <summary>
    ///
    /// </summary>
    void pane(const struct pane &new_pane)
    {
        pane_ = new_pane;
    }

    /// <summary>
    ///
    /// </summary>
    bool has_selections() const
    {
        return !selections_.empty();
    }

    /// <summary>
    ///
    /// </summary>
    void add_selection(const class selection &new_selection)
    {
        selections_.push_back(new_selection);
    }

    /// <summary>
    ///
    /// </summary>
    void clear_selections()
    {
        selections_.clear();
    }

    /// <summary>
    ///
    /// </summary>
    std::vector<xlnt::selection> selections() const
    {
        return selections_;
    }

    /// <summary>
    ///
    /// </summary>
    class xlnt::selection &selection(std::size_t index)
    {
        return selections_.at(index);
    }

    /// <summary>
    ///
    /// </summary>
    void show_grid_lines(bool show)
    {
        show_grid_lines_ = show;
    }

    /// <summary>
    ///
    /// </summary>
    bool show_grid_lines() const
    {
        return show_grid_lines_;
    }

    /// <summary>
    ///
    /// </summary>
    void default_grid_color(bool is_default)
    {
        default_grid_color_ = is_default;
    }

    /// <summary>
    ///
    /// </summary>
    bool default_grid_color() const
    {
        return default_grid_color_;
    }

    /// <summary>
    /// Sets the type of this view.
    /// </summary>
    void type(sheet_view_type new_type)
    {
        type_ = new_type;
    }

    /// <summary>
    /// Returns the type of this view.
    /// </summary>
    sheet_view_type type() const
    {
        return type_;
    }

    /// <summary>
    /// Returns true if this view is requal to rhs based on its id, grid lines setting,
    /// default grid color, pane, and selections.
    /// </summary>
    bool operator==(const sheet_view &rhs) const
    {
        return id_ == rhs.id_ && show_grid_lines_ == rhs.show_grid_lines_
            && default_grid_color_ == rhs.default_grid_color_
            && pane_ == rhs.pane_ && selections_ == rhs.selections_;
    }

private:
    /// <summary>
    ///
    /// </summary>
    std::size_t id_ = 0;

    /// <summary>
    ///
    /// </summary>
    bool show_grid_lines_ = true;

    /// <summary>
    ///
    /// </summary>
    bool default_grid_color_ = true;

    /// <summary>
    ///
    /// </summary>
    sheet_view_type type_ = sheet_view_type::normal;

    /// <summary>
    ///
    /// </summary>
    optional<xlnt::pane> pane_;

    /// <summary>
    ///
    /// </summary>
    std::vector<xlnt::selection> selections_;
};

} // namespace xlnt
