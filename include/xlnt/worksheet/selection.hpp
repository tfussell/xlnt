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

#include <xlnt/xlnt_config.hpp>
#include <xlnt/cell/cell_reference.hpp>
#include <xlnt/worksheet/pane.hpp>
#include <xlnt/worksheet/range_reference.hpp>

namespace xlnt {

/// <summary>
/// The selected area of a worksheet.
/// </summary>
class XLNT_API selection
{
public:
    /// <summary>
    /// Returns true if this selection has a defined active cell.
    /// </summary>
    bool has_active_cell() const
    {
        return active_cell_.is_set();
    }

    /// <summary>
    /// Returns the cell reference of the active cell.
    /// </summary>
    cell_reference active_cell() const
    {
        return active_cell_.get();
    }

    /// <summary>
    /// Sets the active cell to that pointed to by ref.
    /// </summary>
    void active_cell(const cell_reference &ref)
    {
        active_cell_ = ref;
    }

    /// <summary>
    /// Returns the range encompassed by this selection.
    /// </summary>
    range_reference sqref() const
    {
        return sqref_;
    }

    /// <summary>
    /// Returns the sheet quadrant of this selection.
    /// </summary>
    pane_corner pane() const
    {
        return pane_;
    }

    /// <summary>
    /// Sets the sheet quadrant of this selection to corner.
    /// </summary>
    void pane(pane_corner corner)
    {
        pane_ = corner;
    }

    /// <summary>
    /// Returns true if this selection is equal to rhs based on its active cell,
    /// sqref, and pane.
    /// </summary>
    bool operator==(const selection &rhs) const
    {
        return active_cell_ == rhs.active_cell_
            && sqref_ == rhs.sqref_
            && pane_ == rhs.pane_;
    }

private:
    /// <summary>
    /// The active cell
    /// </summary>
    optional<cell_reference> active_cell_;

    /// <summary>
    /// The range
    /// </summary>
    range_reference sqref_;

    /// <summary>
    /// The quadrant
    /// </summary>
    pane_corner pane_;
};

} // namespace xlnt
