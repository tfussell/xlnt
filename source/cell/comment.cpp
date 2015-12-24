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
#include <xlnt/cell/cell.hpp>
#include <xlnt/cell/comment.hpp>

#include "detail/comment_impl.hpp"

namespace xlnt {

comment::comment(detail::comment_impl *d) : d_(d)
{
}

comment::comment(cell parent, const std::string &text, const std::string &author) : d_(nullptr)
{
    d_ = parent.get_comment().d_;
    d_->text_ = text;
    d_->author_ = author;
}

comment::comment() : d_(nullptr)
{
}

comment::~comment()
{
}

std::string comment::get_author() const
{
    return d_->author_;
}

std::string comment::get_text() const
{
    return d_->text_;
}

bool comment::operator==(const xlnt::comment &other) const
{
    return d_ == other.d_;
}

} // namespace xlnt
