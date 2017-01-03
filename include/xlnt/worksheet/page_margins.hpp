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

namespace xlnt {

/// <summary>
/// Describes the margins around a worksheet for printing.
/// </summary>
class XLNT_API page_margins
{
public:
    /// <summary>
    ///
    /// </summary>
    page_margins();

    /// <summary>
    ///
    /// </summary>
    double top() const;

    /// <summary>
    ///
    /// </summary>
    void top(double top);

    /// <summary>
    ///
    /// </summary>
    double left() const;

    /// <summary>
    ///
    /// </summary>
    void left(double left);

    /// <summary>
    ///
    /// </summary>
    double bottom() const;

    /// <summary>
    ///
    /// </summary>
    void bottom(double bottom);

    /// <summary>
    ///
    /// </summary>
    double right() const;

    /// <summary>
    ///
    /// </summary>
    void right(double right);

    /// <summary>
    ///
    /// </summary>
    double header() const;

    /// <summary>
    ///
    /// </summary>
    void header(double header);

    /// <summary>
    ///
    /// </summary>
    double footer() const;

    /// <summary>
    ///
    /// </summary>
    void footer(double footer);

private:
    /// <summary>
    ///
    /// </summary>
    double top_ = 1;

    /// <summary>
    ///
    /// </summary>
    double left_ = 0.75;

    /// <summary>
    ///
    /// </summary>
    double bottom_ = 1;

    /// <summary>
    ///
    /// </summary>
    double right_ = 0.75;

    /// <summary>
    ///
    /// </summary>
    double header_ = 0.5;

    /// <summary>
    ///
    /// </summary>
    double footer_ = 0.5;
};

} // namespace xlnt
