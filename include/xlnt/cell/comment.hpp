// Copyright (c) 2014-2016 Thomas Fussell
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
#include <xlnt/cell/formatted_text.hpp>

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
    comment(const formatted_text &text, const std::string &author);

    /// <summary>
    /// Constructs a new comment with the given unformatted text and author.
    /// </summary>
    comment(const std::string &text, const std::string &author);

    /// <summary>
    /// Return the text that will be displayed for this comment.
    /// </summary>
    formatted_text text() const;
    
    /// <summary>
    /// Return the plain text that will be displayed for this comment without formatting information.
    /// </summary>
    std::string plain_text() const;

    /// <summary>
    /// Return the author of this comment.
    /// </summary>
    std::string author() const;

    /// <summary>
    /// Return true if both comments are equivalent.
    /// </summary>
    friend bool operator==(const comment &left, const comment &right);

private:
    formatted_text text_;
    std::string author_;
};

} // namespace xlnt
