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

#include <cstdint>
#include <string>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/utils/optional.hpp>

namespace xlnt {

class alignment;
class border;
class cell;
class fill;
class font;
class number_format;
class protection;

namespace detail {

struct style_impl;
struct stylesheet;
class xlsx_consumer;

} // namespace detail

/// <summary>
/// Describes a style which has a name and can be applied to multiple individual
/// formats. In Excel this is a "Cell Style".
/// </summary>
class XLNT_API style
{
public:
    /// <summary>
    /// Delete zero-argument constructor
    /// </summary>
    style() = delete;

    /// <summary>
    /// Default copy constructor
    /// </summary>
    style(const style &other) = default;

    /// <summary>
    /// Return the name of this style.
    /// </summary>
    std::string name() const;

    /// <summary>
    ///
    /// </summary>
    style name(const std::string &name);

    /// <summary>
    ///
    /// </summary>
    bool hidden() const;

    /// <summary>
    ///
    /// </summary>
    style hidden(bool value);

    /// <summary>
    ///
    /// </summary>
    optional<bool> custom() const;

    /// <summary>
    ///
    /// </summary>
    style custom(bool value);

    /// <summary>
    ///
    /// </summary>
    optional<std::size_t> builtin_id() const;

    /// <summary>
    ///
    /// </summary>
    style builtin_id(std::size_t builtin_id);

    // Formatting components

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
    style alignment(const xlnt::alignment &new_alignment, bool applied = true);

    /// <summary>
    ///
    /// </summary>
    bool alignment_applied() const;

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
    style border(const xlnt::border &new_border, bool applied = true);

    /// <summary>
    ///
    /// </summary>
    bool border_applied() const;

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
    style fill(const xlnt::fill &new_fill, bool applied = true);

    /// <summary>
    ///
    /// </summary>
    bool fill_applied() const;

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
    style font(const xlnt::font &new_font, bool applied = true);

    /// <summary>
    ///
    /// </summary>
    bool font_applied() const;

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
    style number_format(const xlnt::number_format &new_number_format, bool applied = true);

    /// <summary>
    ///
    /// </summary>
    bool number_format_applied() const;

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
    style protection(const xlnt::protection &new_protection, bool applied = true);

    /// <summary>
    ///
    /// </summary>
    bool protection_applied() const;

    /// <summary>
    ///
    /// </summary>
    bool operator==(const style &other) const;

private:
    friend struct detail::stylesheet;
    friend class detail::xlsx_consumer;

    /// <summary>
    ///
    /// </summary>
    style(detail::style_impl *d);

    /// <summary>
    ///
    /// </summary>
    detail::style_impl *d_;
};

} // namespace xlnt
