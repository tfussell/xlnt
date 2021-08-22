// Copyright (c) 2014-2021 Thomas Fussell
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
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE
//
// @license: http://www.opensource.org/licenses/mit-license.php
// @author: see AUTHORS file

#pragma once

#include <xlnt/xlnt_config.hpp>

namespace xlnt {

/// <summary>
/// Describes the margins around a worksheet for printing.
/// </summary>
class XLNT_API page_margins
{
public:
    /// <summary>
    /// Constructs a page margins objects with Excel-default margins.
    /// </summary>
    page_margins();

    /// <summary>
    /// Returns the top margin
    /// </summary>
    double top() const;

    /// <summary>
    /// Sets the top margin to top
    /// </summary>
    void top(double top);

    /// <summary>
    /// Returns the left margin
    /// </summary>
    double left() const;

    /// <summary>
    /// Sets the left margin to left
    /// </summary>
    void left(double left);

    /// <summary>
    /// Returns the bottom margin
    /// </summary>
    double bottom() const;

    /// <summary>
    /// Sets the bottom margin to bottom
    /// </summary>
    void bottom(double bottom);

    /// <summary>
    /// Returns the right margin
    /// </summary>
    double right() const;

    /// <summary>
    /// Sets the right margin to right
    /// </summary>
    void right(double right);

    /// <summary>
    /// Returns the header margin
    /// </summary>
    double header() const;

    /// <summary>
    /// Sets the header margin to header
    /// </summary>
    void header(double header);

    /// <summary>
    /// Returns the footer margin
    /// </summary>
    double footer() const;

    /// <summary>
    /// Sets the footer margin to footer
    /// </summary>
    void footer(double footer);

    bool operator==(const page_margins &rhs) const;

private:
    /// <summary>
    /// The top margin
    /// </summary>
    double top_ = 1;

    /// <summary>
    /// The left margin
    /// </summary>
    double left_ = 0.75;

    /// <summary>
    /// The bottom margin
    /// </summary>
    double bottom_ = 1;

    /// <summary>
    /// The right margin
    /// </summary>
    double right_ = 0.75;

    /// <summary>
    /// The header margin
    /// </summary>
    double header_ = 0.5;

    /// <summary>
    /// The footer margin
    /// </summary>
    double footer_ = 0.5;
};

} // namespace xlnt
