// Copyright (c) 2014-2021 Thomas Fussell
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
#include <xlnt/cell/rich_text.hpp>

namespace xlnt {

/// <summary>
/// A comment can be applied to a cell to provide extra information about its contents.
/// </summary>
class XLNT_API comment
{
public:
    /// <summary>
    /// Constructs a new blank comment.
    /// </summary>
    comment();

    /// <summary>
    /// Constructs a new comment with the given text and author.
    /// </summary>
    comment(const rich_text &text, const std::string &author);

    /// <summary>
    /// Constructs a new comment with the given unformatted text and author.
    /// </summary>
    comment(const std::string &text, const std::string &author);

    /// <summary>
    /// Returns the text that will be displayed for this comment.
    /// </summary>
    rich_text text() const;

    /// <summary>
    /// Returns the plain text that will be displayed for this comment without formatting information.
    /// </summary>
    std::string plain_text() const;

    /// <summary>
    /// Returns the author of this comment.
    /// </summary>
    std::string author() const;

    /// <summary>
    /// Makes this comment only visible when the associated cell is hovered.
    /// </summary>
    void hide();

    /// <summary>
    /// Makes this comment always visible.
    /// </summary>
    void show();

    /// <summary>
    /// Returns true if this comment is not hidden.
    /// </summary>
    bool visible() const;

    /// <summary>
    /// Sets the absolute position of this cell to the given coordinates.
    /// </summary>
    void position(int left, int top);

    /// <summary>
    /// Returns the distance from the left side of the sheet to the left side of the comment.
    /// </summary>
    int left() const;

    /// <summary>
    /// Returns the distance from the top of the sheet to the top of the comment.
    /// </summary>
    int top() const;

    /// <summary>
    /// Sets the size of the comment.
    /// </summary>
    void size(int width, int height);

    /// <summary>
    /// Returns the width of this comment.
    /// </summary>
    int width() const;

    /// <summary>
    /// Returns the height of this comment.
    /// </summary>
    int height() const;

    /// <summary>
    /// Return true if this comment is equivalent to other.
    /// </summary>
    bool operator==(const comment &other) const;

    /// <summary>
    /// Returns true if this comment is not equivalent to other.
    /// </summary>
    bool operator!=(const comment &other) const;

private:
    /// <summary>
    /// The formatted textual content in this cell displayed directly after the author.
    /// </summary>
    rich_text text_;

    /// <summary>
    /// The name of the person that created this comment.
    /// </summary>
    std::string author_;

    /// <summary>
    /// True if this comment is not hidden.
    /// </summary>
    bool visible_ = false;

    /// <summary>
    /// The fill color
    /// </summary>
    std::string fill_;

    /// <summary>
    /// Distance from the left side of the sheet.
    /// </summary>
    int left_ = 0;

    /// <summary>
    /// Distance from the top of the sheet.
    /// </summary>
    int top_ = 0;

    /// <summary>
    /// Width of the comment box.
    /// </summary>
    int width_ = 200;

    /// <summary>
    /// Height of the comment box.
    /// </summary>
    int height_ = 100;
};

} // namespace xlnt
