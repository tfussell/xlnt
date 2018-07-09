// Copyright (c) 2014-2018 Thomas Fussell
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

namespace xlnt {

/// <summary>
/// The orientation of the worksheet when it is printed.
/// </summary>
enum class XLNT_API orientation
{
    default_orientation,
    portrait,
    landscape
};

/// <summary>
/// The types of page breaks.
/// </summary>
enum class XLNT_API page_break
{
    none = 0,
    row = 1,
    column = 2
};

/// <summary>
/// The possible paper sizes for printing.
/// </summary>
enum class XLNT_API paper_size
{
    letter = 1,
    letter_small = 2,
    tabloid = 3,
    ledger = 4,
    legal = 5,
    statement = 6,
    executive = 7,
    a3 = 8,
    a4 = 9,
    a4_small = 10,
    a5 = 11
};

/// <summary>
/// Defines how a worksheet appears in the workbook.
/// A workbook must have at least one sheet which is visible at all times.
/// </summary>
enum class XLNT_API sheet_state
{
    visible,
    hidden,
    very_hidden
};

/// <summary>
/// Describes how a worksheet will be converted into a page during printing.
/// </summary>
struct XLNT_API page_setup
{
public:
    /// <summary>
    /// Default constructor.
    /// </summary>
    page_setup();

    /// <summary>
    /// Returns the page break.
    /// </summary>
    xlnt::page_break page_break() const;

    /// <summary>
    /// Sets the page break to b.
    /// </summary>
    void page_break(xlnt::page_break b);

    /// <summary>
    /// Returns the current sheet state of this page setup.
    /// </summary>
    xlnt::sheet_state sheet_state() const;

    /// <summary>
    /// Sets the sheet state to sheet_state.
    /// </summary>
    void sheet_state(xlnt::sheet_state sheet_state);

    /// <summary>
    /// Returns the paper size which should be used to print the worksheet using this page setup.
    /// </summary>
    xlnt::paper_size paper_size() const;

    /// <summary>
    /// Sets the paper size of this page setup.
    /// </summary>
    void paper_size(xlnt::paper_size paper_size);

    /// <summary>
    /// Returns true if this worksheet should be scaled to fit on a single page during printing.
    /// </summary>
    bool fit_to_page() const;

    /// <summary>
    /// If true, forces the worksheet to be scaled to fit on a single page during printing.
    /// </summary>
    void fit_to_page(bool fit_to_page);

    /// <summary>
    /// Returns true if the height of this worksheet should be scaled to fit on a printed page.
    /// </summary>
    bool fit_to_height() const;

    /// <summary>
    /// Sets whether the height of the page should be scaled to fit on a printed page.
    /// </summary>
    void fit_to_height(bool fit_to_height);

    /// <summary>
    /// Returns true if the width of this worksheet should be scaled to fit on a printed page.
    /// </summary>
    bool fit_to_width() const;

    /// <summary>
    /// Sets whether the width of the page should be scaled to fit on a printed page.
    /// </summary>
    void fit_to_width(bool fit_to_width);

    /// <summary>
    /// Sets the factor by which the page should be scaled during printing.
    /// </summary>
    void scale(double scale);

    /// <summary>
    /// Returns the factor by which the page should be scaled during printing.
    /// </summary>
    double scale() const;

    /// <summary>
    /// The orientation
    /// </summary>
    xlnt::optional<xlnt::orientation> orientation_;
    /// <summary>
    /// The horizontal dpi
    /// </summary>
    xlnt::optional<std::size_t> horizontal_dpi_;
    /// <summary>
    /// The vertical dpi
    /// </summary>
    xlnt::optional<std::size_t> vertical_dpi_;

    bool operator==(const page_setup &rhs) const;

private:
    /// <summary>
    /// The break
    /// </summary>
    xlnt::page_break break_;

    /// <summary>
    /// The sheet state
    /// </summary>
    xlnt::sheet_state sheet_state_;

    /// <summary>
    /// The paper size
    /// </summary>
    xlnt::paper_size paper_size_;

    /// <summary>
    /// Whether or not to fit to page
    /// </summary>
    bool fit_to_page_;

    /// <summary>
    /// Whether or not to fit to height
    /// </summary>
    bool fit_to_height_;

    /// <summary>
    /// Whether or not to fit to width
    /// </summary>
    bool fit_to_width_;

    /// <summary>
    /// The amount to scale the worksheet
    /// </summary>
    double scale_;
};

} // namespace xlnt