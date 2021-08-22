// Copyright (c) 2014-2021 Thomas Fussell
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
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
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
    /// Text can be underlined in the enumerated ways
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
    /// Constructs a default font. Calibri, size 12
    /// </summary>
    font();

    /// <summary>
    /// Sets the bold state of the font to bold and returns a reference to the font.
    /// </summary>
    font &bold(bool bold);

    /// <summary>
    /// Returns the bold state of the font.
    /// </summary>
    bool bold() const;

    /// <summary>
    /// Sets the vertical alignment of the font to subscript and returns a reference to the font.
    /// </summary>
    font &subscript(bool value);

    /// <summary>
    /// Returns true if this font has a vertical alignment of subscript.
    /// </summary>
    bool subscript() const;

    /// <summary>
    /// Sets the vertical alignment of the font to superscript and returns a reference to the font.
    /// </summary>
    font &superscript(bool value);

    /// <summary>
    /// Returns true if this font has a vertical alignment of superscript.
    /// </summary>
    bool superscript() const;

    /// <summary>
    /// Sets the bold state of the font to bold and returns a reference to the font.
    /// </summary>
    font &italic(bool italic);

    /// <summary>
    /// Returns true if this font has italics applied.
    /// </summary>
    bool italic() const;

    /// <summary>
    /// Sets the bold state of the font to bold and returns a reference to the font.
    /// </summary>
    font &strikethrough(bool strikethrough);

    /// <summary>
    /// Returns true if this font has a strikethrough applied.
    /// </summary>
    bool strikethrough() const;

    /// <summary>
    /// Sets the bold state of the font to bold and returns a reference to the font.
    /// </summary>
    font &outline(bool outline);

    /// <summary>
    /// Returns true if this font has an outline applied.
    /// </summary>
    bool outline() const;

    /// <summary>
    /// Sets the shadow state of the font to shadow and returns a reference to the font.
    /// </summary>
    font &shadow(bool shadow);

    /// <summary>
    /// Returns true if this font has a shadow applied.
    /// </summary>
    bool shadow() const;

    /// <summary>
    /// Sets the underline state of the font to new_underline and returns a reference to the font.
    /// </summary>
    font &underline(underline_style new_underline);

    /// <summary>
    /// Returns true if this font has any type of underline applied.
    /// </summary>
    bool underlined() const;

    /// <summary>
    /// Returns the particular style of underline this font has applied.
    /// </summary>
    underline_style underline() const;

    /// <summary>
    /// Returns true if this font has a defined size.
    /// </summary>
    bool has_size() const;

    /// <summary>
    /// Sets the size of the font to size and returns a reference to the font.
    /// </summary>
    font &size(double size);

    /// <summary>
    /// Returns the size of the font.
    /// </summary>
    double size() const;

    /// <summary>
    /// Returns true if this font has a particular face applied (e.g. "Comic Sans").
    /// </summary>
    bool has_name() const;

    /// <summary>
    /// Sets the font face to name and returns a reference to the font.
    /// </summary>
    font &name(const std::string &name);

    /// <summary>
    /// Returns the name of the font face.
    /// </summary>
    const std::string &name() const;

    /// <summary>
    /// Returns true if this font has a color applied.
    /// </summary>
    bool has_color() const;

    /// <summary>
    /// Sets the color of the font to c and returns a reference to the font.
    /// </summary>
    font &color(const color &c);

    /// <summary>
    /// Returns the color that this font is using.
    /// </summary>
    xlnt::color color() const;

    /// <summary>
    /// Returns true if this font has a family defined.
    /// </summary>
    bool has_family() const;

    /// <summary>
    /// Sets the family index of the font to family and returns a reference to the font.
    /// </summary>
    font &family(std::size_t family);

    /// <summary>
    /// Returns the family index for the font.
    /// </summary>
    std::size_t family() const;

    /// <summary>
    /// Returns true if this font has a charset defined.
    /// </summary>
    bool has_charset() const;

    // TODO: charset should be an enum, not a number

    /// <summary>
    /// Sets the charset of the font to charset and returns a reference to the font.
    /// </summary>
    font &charset(std::size_t charset);

    /// <summary>
    /// Returns the charset of the font.
    /// </summary>
    std::size_t charset() const;

    /// <summary>
    /// Returns true if this font has a scheme.
    /// </summary>
    bool has_scheme() const;

    /// <summary>
    /// Sets the scheme of the font to scheme and returns a reference to the font.
    /// </summary>
    font &scheme(const std::string &scheme);

    /// <summary>
    /// Returns the scheme of this font.
    /// </summary>
    const std::string &scheme() const;

    /// <summary>
    /// Returns true if left is exactly equal to right.
    /// </summary>
    bool operator==(const font &other) const;

    /// <summary>
    /// Returns true if left is not exactly equal to right.
    /// </summary>
    bool operator!=(const font &other) const
    {
        return !operator==(other);
    }

private:
    friend class style;

    /// <summary>
    /// The name of the font
    /// </summary>
    optional<std::string> name_;

    /// <summary>
    /// size
    /// </summary>
    optional<double> size_;

    /// <summary>
    /// bold
    /// </summary>
    bool bold_ = false;

    /// <summary>
    /// italic
    /// </summary>
    bool italic_ = false;

    /// <summary>
    /// superscript
    /// </summary>
    bool superscript_ = false;

    /// <summary>
    /// subscript
    /// </summary>
    bool subscript_ = false;

    /// <summary>
    /// strikethrough
    /// </summary>
    bool strikethrough_ = false;

    /// <summary>
    /// outline
    /// </summary>
    bool outline_ = false;

    /// <summary>
    /// shadow
    /// </summary>
    bool shadow_ = false;

    /// <summary>
    /// underline style
    /// </summary>
    underline_style underline_ = underline_style::none;

    /// <summary>
    /// color
    /// </summary>
    optional<xlnt::color> color_;

    /// <summary>
    /// family
    /// </summary>
    optional<std::size_t> family_;

    /// <summary>
    /// charset
    /// </summary>
    optional<std::size_t> charset_;

    /// <summary>
    /// scheme
    /// </summary>
    optional<std::string> scheme_;
};

} // namespace xlnt
