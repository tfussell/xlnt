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

#include <string>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/styles/color.hpp>
#include <xlnt/utils/optional.hpp>

namespace xlnt {

class style;

/// <summary>
/// Describes the font style of a particular cell.
/// </summary>
class XLNT_API font
{
public:
    /// <summary>
    ///
    /// </summary>
    enum class underline_style
    {
        none,
        double_,
        double_accounting,
        single,
        single_accounting
    };

    /// <summary>
    ///
    /// </summary>
    font();

    /// <summary>
    ///
    /// </summary>
    font &bold(bool bold);

    /// <summary>
    ///
    /// </summary>
    bool bold() const;

    /// <summary>
    ///
    /// </summary>
    font &superscript(bool superscript);

    /// <summary>
    ///
    /// </summary>
    bool superscript() const;

    /// <summary>
    ///
    /// </summary>
    font &italic(bool italic);

    /// <summary>
    ///
    /// </summary>
    bool italic() const;

    /// <summary>
    ///
    /// </summary>
    font &strikethrough(bool strikethrough);

    /// <summary>
    ///
    /// </summary>
    bool strikethrough() const;

    /// <summary>
    ///
    /// </summary>
    font &underline(underline_style new_underline);

    /// <summary>
    ///
    /// </summary>
    bool underlined() const;

    /// <summary>
    ///
    /// </summary>
    underline_style underline() const;

    /// <summary>
    ///
    /// </summary>
    bool has_size() const;

    /// <summary>
    ///
    /// </summary>
    font &size(double size);

    /// <summary>
    ///
    /// </summary>
    double size() const;

    /// <summary>
    ///
    /// </summary>
    bool has_name() const;

    /// <summary>
    ///
    /// </summary>
    font &name(const std::string &name);

    /// <summary>
    ///
    /// </summary>
    std::string name() const;

    /// <summary>
    ///
    /// </summary>
    bool has_color() const;

    /// <summary>
    ///
    /// </summary>
    font &color(const color &c);

    /// <summary>
    ///
    /// </summary>
    xlnt::color color() const;

    /// <summary>
    ///
    /// </summary>
    bool has_family() const;

    /// <summary>
    ///
    /// </summary>
    font &family(std::size_t family);

    /// <summary>
    ///
    /// </summary>
    std::size_t family() const;

    /// <summary>
    ///
    /// </summary>
    bool has_scheme() const;

    /// <summary>
    ///
    /// </summary>
    font &scheme(const std::string &scheme);

    /// <summary>
    ///
    /// </summary>
    std::string scheme() const;

    /// <summary>
    /// Returns true if left is exactly equal to right.
    /// </summary>
    XLNT_API friend bool operator==(const font &left, const font &right);

    /// <summary>
    /// Returns true if left is not exactly equal to right.
    /// </summary>
    XLNT_API friend bool operator!=(const font &left, const font &right)
    {
        return !(left == right);
    }

private:
    friend class style;

    /// <summary>
    ///
    /// </summary>
    optional<std::string> name_;

    /// <summary>
    ///
    /// </summary>
    optional<double> size_;

    /// <summary>
    ///
    /// </summary>
    bool bold_ = false;

    /// <summary>
    ///
    /// </summary>
    bool italic_ = false;

    /// <summary>
    ///
    /// </summary>
    bool superscript_ = false;

    /// <summary>
    ///
    /// </summary>
    bool subscript_ = false;

    /// <summary>
    ///
    /// </summary>
    bool strikethrough_ = false;

    /// <summary>
    ///
    /// </summary>
    underline_style underline_ = underline_style::none;

    /// <summary>
    ///
    /// </summary>
    optional<xlnt::color> color_;

    /// <summary>
    ///
    /// </summary>
    optional<std::size_t> family_;

    /// <summary>
    ///
    /// </summary>
    optional<std::string> scheme_;
};

} // namespace xlnt
