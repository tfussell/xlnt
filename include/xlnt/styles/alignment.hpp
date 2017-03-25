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

#include <xlnt/xlnt_config.hpp>
#include <xlnt/utils/optional.hpp>

namespace xlnt {

/// <summary>
/// Text can be aligned horizontally within a cell in these enumerated ways.
/// </summary>
enum class XLNT_API horizontal_alignment
{
    general,
    left,
    center,
    right,
    fill,
    justify,
    center_continuous,
    distributed
};

/// <summary>
/// Text can be aligned vertically within a cell in these enumerated ways.
/// </summary>
enum class XLNT_API vertical_alignment
{
    top,
    center,
    bottom,
    justify,
    distributed
};

/// <summary>
/// Alignment options that determine how text should be displayed within a cell.
/// </summary>
class XLNT_API alignment
{
public:
    /// <summary>
    /// Returns true if shrink-to-fit has been enabled.
    /// </summary>
    bool shrink() const;

    /// <summary>
    /// Sets whether the font size should be reduced until all of the text fits in a cell without wrapping.
    /// </summary>
    alignment &shrink(bool shrink_to_fit);

    /// <summary>
    /// Returns true if text-wrapping has been enabled.
    /// </summary>
    bool wrap() const;

    /// <summary>
    /// Sets whether text in a cell should continue to multiple lines if it doesn't fit in one line.
    /// </summary>
    alignment &wrap(bool wrap_text);

    /// <summary>
    /// Returns the optional value of indentation width in number of spaces.
    /// </summary>
    optional<int> indent() const;

    /// <summary>
    /// Sets the indent size in number of spaces from the side of the cell. This will only
	/// take effect when left or right horizontal alignment has also been set.
    /// </summary>
    alignment &indent(int indent_size);

    /// <summary>
    /// Returns the optional value of rotation for text in the cell in degrees.
    /// </summary>
    optional<int> rotation() const;

    /// <summary>
    /// Sets the rotation for text in the cell in degrees.
    /// </summary>
    alignment &rotation(int text_rotation);

    /// <summary>
    /// Returns the optional horizontal alignment.
    /// </summary>
    optional<horizontal_alignment> horizontal() const;

    /// <summary>
    /// Sets the horizontal alignment.
    /// </summary>
    alignment &horizontal(horizontal_alignment horizontal);

    /// <summary>
    /// Returns the optional vertical alignment.
    /// </summary>
    optional<vertical_alignment> vertical() const;

    /// <summary>
    /// Sets the vertical alignment.
    /// </summary>
    alignment &vertical(vertical_alignment vertical);

    /// <summary>
    /// Returns true if this alignment is equivalent to other.
    /// </summary>
    bool operator==(const alignment &other) const;

    /// <summary>
    /// Returns true if this alignment is not equivalent to other.
    /// </summary>
    bool operator!=(const alignment &other) const;

private:
	/// <summary>
	/// Whether or not to shrink font size until it fits on one line
	/// </summary>
	bool shrink_to_fit_ = false;

	/// <summary>
	/// Whether or not to wrap text to the next line
	/// </summary>
	bool wrap_text_ = false;

	/// <summary>
	/// The indent in number of spaces from the side
	/// </summary>
	optional<int> indent_;

	/// <summary>
	/// The text roation in degrees
	/// </summary>
    optional<int> text_rotation_;

	/// <summary>
	/// The horizontal alignment
	/// </summary>
    optional<horizontal_alignment> horizontal_;

	/// <summary>
	/// The vertical alignment
	/// </summary>
    optional<vertical_alignment> vertical_;
};

} // namespace xlnt
