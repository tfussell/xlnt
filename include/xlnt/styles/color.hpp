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
///
/// </summary>
class XLNT_API indexed_color
{
public:
    /// <summary>
    ///
    /// </summary>
    indexed_color(std::size_t index);

    /// <summary>
    ///
    /// </summary>
    std::size_t index() const;

    /// <summary>
    ///
    /// </summary>
    void index(std::size_t index);

private:
    /// <summary>
    ///
    /// </summary>
    std::size_t index_;
};

/// <summary>
///
/// </summary>
class XLNT_API theme_color
{
public:
    /// <summary>
    ///
    /// </summary>
    theme_color(std::size_t index);

    /// <summary>
    ///
    /// </summary>
    std::size_t index() const;

    /// <summary>
    ///
    /// </summary>
    void index(std::size_t index);

private:
    /// <summary>
    ///
    /// </summary>
    std::size_t index_;
};

/// <summary>
///
/// </summary>
class XLNT_API rgb_color
{
public:
    /// <summary>
    ///
    /// </summary>
    rgb_color(const std::string &hex_string);

    /// <summary>
    ///
    /// </summary>
    rgb_color(std::uint8_t r, std::uint8_t g, std::uint8_t b, std::uint8_t a = 255);

    /// <summary>
    ///
    /// </summary>
    std::string hex_string() const;

    /// <summary>
    ///
    /// </summary>
    std::uint8_t red() const;

    /// <summary>
    ///
    /// </summary>
    std::uint8_t green() const;

    /// <summary>
    ///
    /// </summary>
    std::uint8_t blue() const;

    /// <summary>
    ///
    /// </summary>
    std::uint8_t alpha() const;

    /// <summary>
    ///
    /// </summary>
    std::array<std::uint8_t, 3> rgb() const;

    /// <summary>
    ///
    /// </summary>
    std::array<std::uint8_t, 4> rgba() const;

private:
    /// <summary>
    ///
    /// </summary>
    static std::array<std::uint8_t, 4> decode_hex_string(const std::string &hex_string);

    /// <summary>
    ///
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
    ///
    /// </summary>
    static const color black();

    /// <summary>
    ///
    /// </summary>
    static const color white();

    /// <summary>
    ///
    /// </summary>
    static const color red();

    /// <summary>
    ///
    /// </summary>
    static const color darkred();

    /// <summary>
    ///
    /// </summary>
    static const color blue();

    /// <summary>
    ///
    /// </summary>
    static const color darkblue();

    /// <summary>
    ///
    /// </summary>
    static const color green();

    /// <summary>
    ///
    /// </summary>
    static const color darkgreen();

    /// <summary>
    ///
    /// </summary>
    static const color yellow();

    /// <summary>
    ///
    /// </summary>
    static const color darkyellow();

    /// <summary>
    ///
    /// </summary>
    color();

    /// <summary>
    ///
    /// </summary>
    color(const rgb_color &rgb);

    /// <summary>
    ///
    /// </summary>
    color(const indexed_color &indexed);

    /// <summary>
    ///
    /// </summary>
    color(const theme_color &theme);

    /// <summary>
    ///
    /// </summary>
    color_type type() const;

    /// <summary>
    ///
    /// </summary>
    bool is_auto() const;

    /// <summary>
    ///
    /// </summary>
    void auto_(bool value);

    /// <summary>
    ///
    /// </summary>
    const rgb_color &rgb() const;

    /// <summary>
    ///
    /// </summary>
    const indexed_color &indexed() const;

    /// <summary>
    ///
    /// </summary>
    const theme_color &theme() const;

    /// <summary>
    ///
    /// </summary>
    double tint() const;

    /// <summary>
    ///
    /// </summary>
    void tint(double tint);

    /// <summary>
    ///
    /// </summary>
    bool operator==(const color &other) const;

    /// <summary>
    ///
    /// </summary>
    bool operator!=(const color &other) const
    {
        return !(*this == other);
    }

private:
    /// <summary>
    ///
    /// </summary>
    void assert_type(color_type t) const;

    /// <summary>
    ///
    /// </summary>
    color_type type_;

    /// <summary>
    ///
    /// </summary>
    rgb_color rgb_;

    /// <summary>
    ///
    /// </summary>
    indexed_color indexed_;

    /// <summary>
    ///
    /// </summary>
    theme_color theme_;

    /// <summary>
    ///
    /// </summary>
    double tint_;

    /// <summary>
    ///
    /// </summary>
    bool auto__;
};

} // namespace xlnt
