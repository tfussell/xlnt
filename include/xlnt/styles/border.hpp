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
#include <xlnt/styles/color.hpp>
#include <xlnt/utils/optional.hpp>

namespace xlnt {

/// <summary>
/// Enumerates the sides of a cell to which a border style can be applied.
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

/// <summary>
/// Enumerates the pattern of the border lines on a particular side.
/// </summary>
enum class XLNT_API border_style
{
    none,
    dashdot,
    dashdotdot,
    dashed,
    dotted,
    double_,
    hair,
    medium,
    mediumdashdot,
    mediumdashdotdot,
    mediumdashed,
    slantdashdot,
    thick,
    thin
};

/// <summary>
/// Cells can have borders that go from the top-left to bottom-right
/// or from the top-right to bottom-left, or both, or neither.
/// Used by style->border.
/// </summary>
enum class XLNT_API diagonal_direction
{
    neither,
    up,
    down,
    both
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
    /// Each side of a cell can have a border_property applied to it to change
    /// how it is displayed.
    /// </summary>
    class XLNT_API border_property
    {
    public:
        /// <summary>
        /// Returns the color of the side.
        /// </summary>
        optional<class color> color() const;

        /// <summary>
        /// Sets the color of the side and returns a reference to the side properties.
        /// </summary>
        border_property &color(const xlnt::color &c);

        /// <summary>
        /// Returns the style of the side.
        /// </summary>
        optional<border_style> style() const;

        /// <summary>
        /// Sets the style of the side and returns a reference to the side properties.
        /// </summary>
        border_property &style(border_style style);

        /// <summary>
        /// Returns true if left is exactly equal to right.
        /// </summary>
        bool operator==(const border_property &right) const;

        /// <summary>
        /// Returns true if left is not exactly equal to right.
        /// </summary>
        bool operator!=(const border_property &right) const;

    private:
        /// <summary>
        /// The color of the side
        /// </summary>
        optional<class color> color_;

        /// <summary>
        /// The style of the side
        /// </summary>
        optional<border_style> style_;
    };

    /// <summary>
    /// Returns a vector containing all of the border sides to be used for iteration.
    /// </summary>
    static const std::vector<border_side> &all_sides();

    /// <summary>
    /// Constructs a default border.
    /// </summary>
    border();

    /// <summary>
    /// Returns the border properties of the given side.
    /// </summary>
    optional<border_property> side(border_side s) const;

    /// <summary>
    /// Sets the border properties of the side s to prop.
    /// </summary>
    border &side(border_side s, const border_property &prop);

    /// <summary>
    /// Returns the diagonal direction of this border.
    /// </summary>
    optional<diagonal_direction> diagonal() const;

    /// <summary>
    /// Sets the diagonal direction of this border to dir.
    /// </summary>
    border &diagonal(diagonal_direction dir);

    /// <summary>
    /// Returns true if left is exactly equal to right.
    /// </summary>
    bool operator==(const border &right) const;

    /// <summary>
    /// Returns true if left is not exactly equal to right.
    /// </summary>
    bool operator!=(const border &right) const;

private:
    /// <summary>
    /// Start side (i.e. left) border properties
    /// </summary>
    optional<border_property> start_;

    /// <summary>
    /// End side (i.e. right) border properties
    /// </summary>
    optional<border_property> end_;

    /// <summary>
    /// Top side border properties
    /// </summary>
    optional<border_property> top_;

    /// <summary>
    /// Bottom side border properties
    /// </summary>
    optional<border_property> bottom_;

    /// <summary>
    /// Vertical border properties
    /// </summary>
    optional<border_property> vertical_;

    /// <summary>
    /// Horizontal border properties
    /// </summary>
    optional<border_property> horizontal_;

    /// <summary>
    /// Diagonal border properties
    /// </summary>
    optional<border_property> diagonal_;

    /// <summary>
    /// Direction of diagonal border properties to be applied
    /// </summary>
    optional<diagonal_direction> diagonal_direction_;
};

} // namespace xlnt
