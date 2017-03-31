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
#include <xlnt/cell/comment.hpp>

namespace xlnt {

comment::comment()
    : comment("", "")
{
}

comment::comment(const rich_text &text, const std::string &author)
    : text_(text), author_(author)
{
}

comment::comment(const std::string &text, const std::string &author)
    : text_(), author_(author)
{
    text_.plain_text(text);
}

rich_text comment::text() const
{
    return text_;
}

std::string comment::plain_text() const
{
    return text_.plain_text();
}

std::string comment::author() const
{
    return author_;
}

void comment::hide()
{
    visible_ = false;
}

void comment::show()
{
    visible_ = true;
}

void comment::position(int left, int top)
{
    left_ = left;
    top_ = top;
}

void comment::size(int width, int height)
{
    width_ = width;
    height_ = height;
}

bool comment::visible() const
{
    return visible_;
}

int comment::left() const
{
    return left_;
}

int comment::top() const
{
    return top_;
}

int comment::width() const
{
    return width_;
}

int comment::height() const
{
    return height_;
}

bool comment::operator==(const comment &other) const
{
    return text_ == other.text_ && author_ == other.author_;
}

bool comment::operator!=(const comment &other) const
{
    return !(*this == other);
}

} // namespace xlnt
