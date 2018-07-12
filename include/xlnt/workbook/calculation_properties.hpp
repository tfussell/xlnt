// Copyright (c) 2016-2018 Thomas Fussell
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
/// Workbook file properties relating to calculations.
/// </summary>
class XLNT_API calculation_properties
{
public:
    /// <summary>
    /// The version of calculation engine used to calculate cell formula values.
    /// If this is older than the version of the Excel calculation engine opening
    /// the workbook, cell values will be recalculated.
    /// </summary>
    std::size_t calc_id;

    /// <summary>
    /// If this is true, concurrent calculation will be enabled for the workbook.
    /// </summary>
    bool concurrent_calc;
};

inline bool operator==(const calculation_properties &lhs, const calculation_properties &rhs)
{
    return lhs.calc_id == rhs.calc_id
        && lhs.concurrent_calc == rhs.concurrent_calc;
}

} // namespace xlnt
