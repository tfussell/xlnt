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

#include <unordered_map>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/styles/color.hpp>
#include <xlnt/utils/optional.hpp>

namespace xlnt {

/// <summary>
///
/// </summary>
enum class XLNT_API pattern_fill_type
{
    none,
    solid,
    mediumgray,
    darkgray,
    lightgray,
    darkhorizontal,
    darkvertical,
    darkdown,
    darkup,
    darkgrid,
    darktrellis,
    lighthorizontal,
    lightvertical,
    lightdown,
    lightup,
    lightgrid,
    lighttrellis,
    gray125,
    gray0625
};

/// <summary>
///
/// </summary>
class XLNT_API pattern_fill
{
public:
    /// <summary>
    ///
    /// </summary>
    pattern_fill();

    /// <summary>
    ///
    /// </summary>
    pattern_fill_type type() const;

    /// <summary>
    ///
    /// </summary>
    pattern_fill &type(pattern_fill_type new_type);

    /// <summary>
    ///
    /// </summary>
    optional<color> foreground() const;

    /// <summary>
    ///
    /// </summary>
    pattern_fill &foreground(const color &foreground);

    /// <summary>
    ///
    /// </summary>
    optional<color> background() const;

    /// <summary>
    ///
    /// </summary>
    pattern_fill &background(const color &background);

    /// <summary>
    /// Returns true if left is exactly equal to right.
    /// </summary>
    XLNT_API friend bool operator==(const pattern_fill &left, const pattern_fill &right);

    /// <summary>
    /// Returns true if left is not exactly equal to right.
    /// </summary>
    XLNT_API friend bool operator!=(const pattern_fill &left, const pattern_fill &right)
    {
        return !(left == right);
    }

private:
    /// <summary>
    ///
    /// </summary>
    pattern_fill_type type_ = pattern_fill_type::none;

    /// <summary>
    ///
    /// </summary>
    optional<color> foreground_;

    /// <summary>
    ///
    /// </summary>
    optional<color> background_;
};

/// <summary>
///
/// </summary>
enum class XLNT_API gradient_fill_type
{
    linear,
    path
};

/// <summary>
///
/// </summary>
class XLNT_API gradient_fill
{
public:
    /// <summary>
    ///
    /// </summary>
    gradient_fill();

    /// <summary>
    ///
    /// </summary>
    gradient_fill_type type() const;

    // Type

    /// <summary>
    ///
    /// </summary>
    gradient_fill &type(gradient_fill_type new_type);

    // Degree

    /// <summary>
    ///
    /// </summary>
    gradient_fill &degree(double degree);

    /// <summary>
    ///
    /// </summary>
    double degree() const;

    // Left

    /// <summary>
    ///
    /// </summary>
    double left() const;

    /// <summary>
    ///
    /// </summary>
    gradient_fill &left(double value);

    // Right

    /// <summary>
    ///
    /// </summary>
    double right() const;

    /// <summary>
    ///
    /// </summary>
    gradient_fill &right(double value);

    // Top

    /// <summary>
    ///
    /// </summary>
    double top() const;

    /// <summary>
    ///
    /// </summary>
    gradient_fill &top(double value);

    // Bottom

    /// <summary>
    ///
    /// </summary>
    double bottom() const;

    /// <summary>
    ///
    /// </summary>
    gradient_fill &bottom(double value);

    // Stops

    /// <summary>
    ///
    /// </summary>
    gradient_fill &add_stop(double position, color stop_color);

    /// <summary>
    ///
    /// </summary>
    gradient_fill &clear_stops();

    /// <summary>
    ///
    /// </summary>
    std::unordered_map<double, color> stops() const;

    /// <summary>
    /// Returns true if left is exactly equal to right.
    /// </summary>
    XLNT_API friend bool operator==(const gradient_fill &left, const gradient_fill &right);

    /// <summary>
    /// Returns true if left is not exactly equal to right.
    /// </summary>
    XLNT_API friend bool operator!=(const gradient_fill &left, const gradient_fill &right)
    {
        return !(left == right);
    }

private:
    /// <summary>
    ///
    /// </summary>
    gradient_fill_type type_ = gradient_fill_type::linear;

    /// <summary>
    ///
    /// </summary>
    double degree_ = 0;

    /// <summary>
    ///
    /// </summary>
    double left_ = 0;

    /// <summary>
    ///
    /// </summary>
    double right_ = 0;

    /// <summary>
    ///
    /// </summary>
    double top_ = 0;

    /// <summary>
    ///
    /// </summary>
    double bottom_ = 0;

    /// <summary>
    ///
    /// </summary>
    std::unordered_map<double, color> stops_;
};

/// <summary>
///
/// </summary>
enum class XLNT_API fill_type
{
    pattern,
    gradient
};

/// <summary>
/// Describes the fill style of a particular cell.
/// </summary>
class XLNT_API fill
{
public:
    /// <summary>
    /// Helper method for the most common use case, setting the fill color of a cell to a single solid color.
    /// The foreground and background colors of a fill are not the same as the foreground and background colors
    /// of a cell. When setting a fill color in Excel, a new fill is created with the given color as the fill's
    /// fgColor and index color number 64 as the bgColor. This method creates a fill in the same way.
    /// </summary>
    static fill solid(const color &fill_color);

    /// <summary>
    /// Constructs a fill initialized as a none-type pattern fill with no
    /// foreground or background colors.
    /// </summary>
    fill();

    /// <summary>
    /// Constructs a fill initialized as a pattern fill based on the given pattern.
    /// </summary>
    fill(const pattern_fill &pattern);

    /// <summary>
    /// Constructs a fill initialized as a gradient fill based on the given gradient.
    /// </summary>
    fill(const gradient_fill &gradient);

    /// <summary>
    /// Returns the fill_type of this fill depending on how it was constructed.
    /// </summary>
    fill_type type() const;

    /// <summary>
    /// Returns the gradient fill represented by this fill.
    /// Throws an invalid_attribute exception if this is not a gradient fill.
    /// </summary>
    class gradient_fill gradient_fill() const;

    /// <summary>
    /// Returns the pattern fill represented by this fill.
    /// Throws an invalid_attribute exception if this is not a pattern fill.
    /// </summary>
    class pattern_fill pattern_fill() const;

    /// <summary>
    /// Returns true if left is exactly equal to right.
    /// </summary>
    XLNT_API friend bool operator==(const fill &left, const fill &right);

    /// <summary>
    /// Returns true if left is not exactly equal to right.
    /// </summary>
    XLNT_API friend bool operator!=(const fill &left, const fill &right)
    {
        return !(left == right);
    }

private:
    /// <summary>
    ///
    /// </summary>
    fill_type type_ = fill_type::pattern;

    /// <summary>
    ///
    /// </summary>
    xlnt::gradient_fill gradient_;

    /// <summary>
    ///
    /// </summary>
    xlnt::pattern_fill pattern_;
};

} // namespace xlnt
