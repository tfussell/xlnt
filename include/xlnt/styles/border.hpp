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

#include <cstddef>
#include <functional>
#include <unordered_map>
#include <vector>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/styles/border_style.hpp>
#include <xlnt/styles/color.hpp>
#include <xlnt/styles/diagonal_direction.hpp>
#include <xlnt/utils/optional.hpp>

namespace xlnt {

/// <summary>
///
/// </summary>
enum class XLNT_API border_side
{
    start,
    end,
    top,
    bottom,
    diagonal,
    vertical,
    horizontal
};

} // namespace xlnt

namespace xlnt {

/// <summary>
/// Describes the border style of a particular cell.
/// </summary>
class XLNT_API border
{
public:
    /// <summary>
    ///
    /// </summary>
    class XLNT_API border_property
    {
    public:
        /// <summary>
        ///
        /// </summary>
        optional<class color> color() const;

        /// <summary>
        ///
        /// </summary>
        border_property &color(const xlnt::color &c);

        /// <summary>
        ///
        /// </summary>
        optional<border_style> style() const;

        /// <summary>
        ///
        /// </summary>
        border_property &style(border_style style);

        /// <summary>
        /// Returns true if left is exactly equal to right.
        /// </summary>
        friend bool operator==(const border_property &left, const border_property &right);

        /// <summary>
        /// Returns true if left is not exactly equal to right.
        /// </summary>
        friend bool operator!=(const border_property &left, const border_property &right)
        {
            return !(left == right);
        }

    private:
        /// <summary>
        ///
        /// </summary>
        optional<class color> color_;

        /// <summary>
        ///
        /// </summary>
        optional<border_style> style_;
    };

    /// <summary>
    ///
    /// </summary>
    static const std::vector<border_side> &all_sides();

    /// <summary>
    ///
    /// </summary>
    border();

    /// <summary>
    ///
    /// </summary>
    optional<border_property> side(border_side s) const;

    /// <summary>
    ///
    /// </summary>
    border &side(border_side s, const border_property &prop);

    /// <summary>
    ///
    /// </summary>
    optional<diagonal_direction> diagonal() const;

    /// <summary>
    ///
    /// </summary>
    border &diagonal(diagonal_direction dir);

    /// <summary>
    /// Returns true if left is exactly equal to right.
    /// </summary>
    XLNT_API friend bool operator==(const border &left, const border &right);

    /// <summary>
    /// Returns true if left is not exactly equal to right.
    /// </summary>
    XLNT_API friend bool operator!=(const border &left, const border &right)
    {
        return !(left == right);
    }

private:
    /// <summary>
    ///
    /// </summary>
    optional<border_property> start_;

    /// <summary>
    ///
    /// </summary>
    optional<border_property> end_;

    /// <summary>
    ///
    /// </summary>
    optional<border_property> top_;

    /// <summary>
    ///
    /// </summary>
    optional<border_property> bottom_;

    /// <summary>
    ///
    /// </summary>
    optional<border_property> vertical_;

    /// <summary>
    ///
    /// </summary>
    optional<border_property> horizontal_;

    /// <summary>
    ///
    /// </summary>
    optional<border_property> diagonal_;

    // bool outline_ = true;

    /// <summary>
    ///
    /// </summary>
    optional<diagonal_direction> diagonal_direction_;
};

} // namespace xlnt
