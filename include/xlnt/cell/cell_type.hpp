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

/// <summary>
/// Enumerates the possible types a cell can be determined by it's current value.
/// </summary>

namespace xlnt {

/// <summary>
/// Enumerates the possible types a cell can be determined by it's current value.
/// </summary>
enum class XLNT_API cell_type
{
    /// no value
    empty,
    /// value is TRUE or FALSE
    boolean,
    /// value is an ISO 8601 formatted date
    date,
    /// value is a known error code such as \#VALUE!
    error,
    /// value is a string stored in the cell
    inline_string,
    /// value is a number
    number,
    /// value is a string shared with other cells to save space
    shared_string,
    /// value is the string result of a formula
    formula_string
};

} // namespace xlnt
