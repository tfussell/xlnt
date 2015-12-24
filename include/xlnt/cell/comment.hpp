// Copyright (c) 2014-2016 Thomas Fussell
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

namespace xlnt {

class cell;
namespace detail { struct comment_impl; }

/// <summary>
/// A comment can be applied to a cell to provide extra information.
/// </summary>
class XLNT_CLASS comment
{
  public:
    /// <summary>
    /// The default constructor makes an invalid comment without a parent cell.
    /// </summary>
    comment();
    
    /// <summary>
    /// Constructs a comment applied to the given cell, parent, and with the comment
    /// text and author set to the provided respective values.
    comment(cell parent, const std::string &text, const std::string &auth);
    
    ~comment();
    
    /// <summary>
    /// Return the text that will be displayed for this comment.
    /// </summary>
    std::string get_text() const;
    
    /// <summary>
    /// Return the author of this comment.
    /// </summary>
    std::string get_author() const;

    /// <summary>
    /// True if the comments point to the same sell (false if
    /// they are different cells but identical comments). Note
    /// that a cell can only have one comment and a comment
    /// can only be applied to one cell.
    /// </summary>
    bool operator==(const comment &other) const;

  private:
    friend class cell; // cell needs access to private constructor
    
    /// <summary>
    /// Construct a comment from an implementation of a comment.
    /// </summary>
    comment(detail::comment_impl *d);
    
    /// <summary>
    /// Pointer to the implementation of this comment.
    /// This allows comments to be passed by value while
    /// retaining the ability to modify the parent cell.
    /// </summary>
    detail::comment_impl *d_;
};

} // namespace xlnt
