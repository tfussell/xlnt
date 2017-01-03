// Copyright (c) 2014-2017 Thomas Fussell
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
#include <string>

#include <xlnt/xlnt_config.hpp>

namespace xlnt {

class alignment;
class border;
class cell;
class fill;
class font;
class number_format;
class protection;
class style;

namespace detail {

struct format_impl;
struct stylesheet;

} // namespace detail

/// <summary>
/// Describes the formatting of a particular cell.
/// </summary>
class XLNT_API format
{
public:
    /// <summary>
    ///
    /// </summary>
    std::size_t id() const;

    // Alignment

    /// <summary>
    ///
    /// </summary>
    class alignment &alignment();

    /// <summary>
    ///
    /// </summary>
    const class alignment &alignment() const;

    /// <summary>
    ///
    /// </summary>
    format alignment(const xlnt::alignment &new_alignment, bool applied);

    /// <summary>
    ///
    /// </summary>
    bool alignment_applied() const;

    // Border

    /// <summary>
    ///
    /// </summary>
    class border &border();

    /// <summary>
    ///
    /// </summary>
    const class border &border() const;

    /// <summary>
    ///
    /// </summary>
    format border(const xlnt::border &new_border, bool applied);

    /// <summary>
    ///
    /// </summary>
    bool border_applied() const;

    // Fill

    /// <summary>
    ///
    /// </summary>
    class fill &fill();

    /// <summary>
    ///
    /// </summary>
    const class fill &fill() const;

    /// <summary>
    ///
    /// </summary>
    format fill(const xlnt::fill &new_fill, bool applied);

    /// <summary>
    ///
    /// </summary>
    bool fill_applied() const;

    // Font

    /// <summary>
    ///
    /// </summary>
    class font &font();

    /// <summary>
    ///
    /// </summary>
    const class font &font() const;

    /// <summary>
    ///
    /// </summary>
    format font(const xlnt::font &new_font, bool applied);

    /// <summary>
    ///
    /// </summary>
    bool font_applied() const;

    // Number Format

    /// <summary>
    ///
    /// </summary>
    class number_format &number_format();

    /// <summary>
    ///
    /// </summary>
    const class number_format &number_format() const;

    /// <summary>
    ///
    /// </summary>
    format number_format(const xlnt::number_format &new_number_format, bool applied);

    /// <summary>
    ///
    /// </summary>
    bool number_format_applied() const;

    // Protection

    /// <summary>
    ///
    /// </summary>
    class protection &protection();

    /// <summary>
    ///
    /// </summary>
    const class protection &protection() const;

    /// <summary>
    ///
    /// </summary>
    format protection(const xlnt::protection &new_protection, bool applied);

    /// <summary>
    ///
    /// </summary>
    bool protection_applied() const;

    // Style

    /// <summary>
    ///
    /// </summary>
    bool has_style() const;

    /// <summary>
    ///
    /// </summary>
    void clear_style();

    /// <summary>
    ///
    /// </summary>
    format style(const std::string &name);

    /// <summary>
    ///
    /// </summary>
    format style(const class style &new_style);

    /// <summary>
    ///
    /// </summary>
    const class style style() const;

private:
    friend struct detail::stylesheet;
    friend class cell;

    /// <summary>
    ///
    /// </summary>
    format(detail::format_impl *d);

    /// <summary>
    ///
    /// </summary>
    detail::format_impl *d_;
};

} // namespace xlnt
