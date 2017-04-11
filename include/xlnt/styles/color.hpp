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

#include <array>
#include <string>

#include <xlnt/xlnt_config.hpp>

namespace xlnt {

/// <summary>
/// An indexed color encapsulates a simple index to a color in the indexedColors of the stylesheet.
/// </summary>
class XLNT_API indexed_color
{
public:
    //TODO: should this be explicit?
    /// <summary>
    /// Constructs an indexed_color from an index.
    /// </summary>
    indexed_color(std::size_t index);

    /// <summary>
    /// Returns the index this color points to.
    /// </summary>
    std::size_t index() const;

    /// <summary>
    /// Sets the index to index.
    /// </summary>
    void index(std::size_t index);

private:
    /// <summary>
    /// The index of this color
    /// </summary>
    std::size_t index_;
};

/// <summary>
/// A theme color encapsulates a color derived from the theme.
/// </summary>
class XLNT_API theme_color
{
public:
    /// <summary>
    /// Constructs a theme_color from an index.
    /// </summary>
    theme_color(std::size_t index);

    /// <summary>
    /// Returns the index of the color in the theme this points to.
    /// </summary>
    std::size_t index() const;

    /// <summary>
    /// Sets the index of this color to index.
    /// </summary>
    void index(std::size_t index);

private:
    /// <summary>
    /// The index of the color
    /// </summary>
    std::size_t index_;
};

/// <summary>
/// An RGB color describes a color in terms of its red, green, blue, and alpha components.
/// </summary>
class XLNT_API rgb_color
{
public:
    /// <summary>
    /// Constructs an RGB color from a string in the form \#[aa]rrggbb
    /// </summary>
    rgb_color(const std::string &hex_string);

    /// <summary>
    /// Constructs an RGB color from red, green, and blue values in the range 0 to 255
	/// plus an optional alpha which defaults to fully opaque.
    /// </summary>
    rgb_color(std::uint8_t r, std::uint8_t g, std::uint8_t b, std::uint8_t a = 255);

    /// <summary>
    /// Returns a string representation of this color in the form \#aarrggbb
    /// </summary>
    std::string hex_string() const;

    /// <summary>
    /// Returns a byte representing the red component of this color
    /// </summary>
    std::uint8_t red() const;

    /// <summary>
	/// Returns a byte representing the red component of this color
    /// </summary>
    std::uint8_t green() const;

    /// <summary>
	/// Returns a byte representing the blue component of this color
    /// </summary>
    std::uint8_t blue() const;

    /// <summary>
	/// Returns a byte representing the alpha component of this color
    /// </summary>
    std::uint8_t alpha() const;

    /// <summary>
    /// Returns the red, green, and blue components of this color separately in an array in that order.
    /// </summary>
    std::array<std::uint8_t, 3> rgb() const;

    /// <summary>
	/// Returns the red, green, blue, and alpha components of this color separately in an array in that order.
    /// </summary>
    std::array<std::uint8_t, 4> rgba() const;

private:
    /// <summary>
    /// The four bytes of this color
    /// </summary>
    std::array<std::uint8_t, 4> rgba_;
};

/// <summary>
/// Some colors are references to colors rather than having a particular RGB value.
/// </summary>
enum class color_type
{
    indexed,
    theme,
    rgb
};

/// <summary>
/// Colors can be applied to many parts of a cell's style.
/// </summary>
class XLNT_API color
{
public:
    /// <summary>
    /// Returns the color \#000000
    /// </summary>
    static const color black();

    /// <summary>
    /// Returns the color \#ffffff
    /// </summary>
    static const color white();

    /// <summary>
	/// Returns the color \#ff0000
    /// </summary>
    static const color red();

    /// <summary>
	/// Returns the color \#8b0000
    /// </summary>
    static const color darkred();

    /// <summary>
	/// Returns the color \#00ff00
    /// </summary>
    static const color blue();

    /// <summary>
	/// Returns the color \#008b00
    /// </summary>
    static const color darkblue();

    /// <summary>
	/// Returns the color \#0000ff
    /// </summary>
    static const color green();

    /// <summary>
	/// Returns the color \#00008b
    /// </summary>
    static const color darkgreen();

    /// <summary>
	/// Returns the color \#ffff00
    /// </summary>
    static const color yellow();

    /// <summary>
	/// Returns the color \#cccc00
    /// </summary>
    static const color darkyellow();

    /// <summary>
    /// Constructs a default color
    /// </summary>
    color();

    /// <summary>
    /// Constructs a color from a given RGB color
    /// </summary>
    color(const rgb_color &rgb);

    /// <summary>
    /// Constructs a color from a given indexed color
    /// </summary>
    color(const indexed_color &indexed);

    /// <summary>
    /// Constructs a color from a given theme color
    /// </summary>
    color(const theme_color &theme);

    /// <summary>
    /// Returns the type of this color
    /// </summary>
    color_type type() const;

    /// <summary>
    /// Returns true if this color has been set to auto
    /// </summary>
    bool auto_() const;

    /// <summary>
    /// Sets the auto property of this color to value
    /// </summary>
    void auto_(bool value);

    /// <summary>
	/// Returns the internal indexed color representing this color. If this is not an RGB color,
	/// an invalid_attribute exception will be thrown.
    /// </summary>
    rgb_color rgb() const;

    /// <summary>
    /// Returns the internal indexed color representing this color. If this is not an indexed color,
	/// an invalid_attribute exception will be thrown.
    /// </summary>
    indexed_color indexed() const;

    /// <summary>
	/// Returns the internal indexed color representing this color. If this is not a theme color,
	/// an invalid_attribute exception will be thrown.
    /// </summary>
    theme_color theme() const;

    /// <summary>
    /// Returns the tint of this color.
    /// </summary>
    double tint() const;

    /// <summary>
    /// Sets the tint of this color to tint. Tints lighten or darken an existing color by multiplying the color with the tint.
    /// </summary>
    void tint(double tint);

    /// <summary>
    /// Returns true if this color is equivalent to other
    /// </summary>
    bool operator==(const color &other) const;

    /// <summary>
    /// Returns true if this color is not equivalent to other
    /// </summary>
	bool operator!=(const color &other) const;

private:
    /// <summary>
    /// Throws an invalid_attribute exception if the given type is different from this color's type
    /// </summary>
    void assert_type(color_type t) const;

    /// <summary>
    /// The type of this color
    /// </summary>
    color_type type_;

    /// <summary>
    /// The internal RGB color. Only valid when this color has a type of rgb
    /// </summary>
    rgb_color rgb_;

    /// <summary>
	/// The internal RGB color. Only valid when this color has a type of indexed
    /// </summary>
    indexed_color indexed_;

    /// <summary>
	/// The internal RGB color. Only valid when this color has a type of theme
    /// </summary>
    theme_color theme_;

    /// <summary>
    /// The tint of this color
    /// </summary>
    double tint_ = 0.0;

    /// <summary>
    /// Whether or not this is an auto color
    /// </summary>
    bool auto__ = false;
};

} // namespace xlnt
