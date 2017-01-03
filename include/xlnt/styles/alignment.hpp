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
#include <xlnt/styles/horizontal_alignment.hpp>
#include <xlnt/styles/vertical_alignment.hpp>
#include <xlnt/utils/optional.hpp>

namespace xlnt {

/// <summary>
/// Alignment options for use in cell formats.
/// </summary>
class XLNT_API alignment
{
public:
    /// <summary>
    ///
    /// </summary>
    bool shrink() const;

    /// <summary>
    ///
    /// </summary>
    alignment &shrink(bool shrink_to_fit);

    /// <summary>
    ///
    /// </summary>
    bool wrap() const;

    /// <summary>
    ///
    /// </summary>
    alignment &wrap(bool wrap_text);

    /// <summary>
    ///
    /// </summary>
    optional<int> indent() const;

    /// <summary>
    ///
    /// </summary>
    alignment &indent(int indent_size);

    /// <summary>
    ///
    /// </summary>
    optional<int> rotation() const;

    /// <summary>
    ///
    /// </summary>
    alignment &rotation(int text_rotation);

    /// <summary>
    ///
    /// </summary>
    optional<horizontal_alignment> horizontal() const;

    /// <summary>
    ///
    /// </summary>
    alignment &horizontal(horizontal_alignment horizontal);

    /// <summary>
    ///
    /// </summary>
    optional<vertical_alignment> vertical() const;

    /// <summary>
    ///
    /// </summary>
    alignment &vertical(vertical_alignment vertical);

    /// <summary>
    /// Returns true if left is exactly equal to right.
    /// </summary>
    XLNT_API friend bool operator==(const alignment &left, const alignment &right);

    /// <summary>
    /// Returns true if left is not exactly equal to right.
    /// </summary>
    XLNT_API friend bool operator!=(const alignment &left, const alignment &right)
    {
        return !(left == right);
    }

private:
    bool shrink_to_fit_ = false;
    bool wrap_text_ = false;
    optional<int> indent_;
    optional<int> text_rotation_;
    optional<horizontal_alignment> horizontal_;
    optional<vertical_alignment> vertical_;
};

} // namespace xlnt
