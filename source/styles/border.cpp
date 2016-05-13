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
#include <xlnt/styles/border.hpp>

namespace xlnt {

std::experimental::optional<side> &border::get_start()
{
    return start_;
}

const std::experimental::optional<side> &border::get_start() const
{
    return start_;
}

std::experimental::optional<side> &border::get_end()
{
    return end_;
}

const std::experimental::optional<side> &border::get_end() const
{
    return end_;
}

std::experimental::optional<side> &border::get_left()
{
    return left_;
}

const std::experimental::optional<side> &border::get_left() const
{
    return left_;
}

std::experimental::optional<side> &border::get_right()
{
    return right_;
}

const std::experimental::optional<side> &border::get_right() const
{
    return right_;
}

std::experimental::optional<side> &border::get_top()
{
    return top_;
}

const std::experimental::optional<side> &border::get_top() const
{
    return top_;
}

std::experimental::optional<side> &border::get_bottom()
{
    return bottom_;
}

const std::experimental::optional<side> &border::get_bottom() const
{
    return bottom_;
}

std::experimental::optional<side> &border::get_diagonal()
{
    return diagonal_;
}

const std::experimental::optional<side> &border::get_diagonal() const
{
    return diagonal_;
}

std::experimental::optional<side> &border::get_vertical()
{
    return vertical_;
}

const std::experimental::optional<side> &border::get_vertical() const
{
    return vertical_;
}

std::experimental::optional<side> &border::get_horizontal()
{
    return horizontal_;
}

const std::experimental::optional<side> &border::get_horizontal() const
{
    return horizontal_;
}

std::string border::to_hash_string() const
{
    std::string hash_string;

    hash_string.append(start_ ? "1" : "0");
    if (start_) hash_string.append(std::to_string(start_->hash()));
    
    hash_string.append(end_ ? "1" : "0");
    if (end_) hash_string.append(std::to_string(end_->hash()));
    
    hash_string.append(left_ ? "1" : "0");
    if (left_) hash_string.append(std::to_string(left_->hash()));
    
    hash_string.append(right_ ? "1" : "0");
    if (right_) hash_string.append(std::to_string(right_->hash()));
    
    hash_string.append(top_ ? "1" : "0");
    if (top_) hash_string.append(std::to_string(top_->hash()));
    
    hash_string.append(bottom_ ? "1" : "0");
    if (bottom_) hash_string.append(std::to_string(bottom_->hash()));
    
    hash_string.append(diagonal_ ? "1" : "0");
    if (diagonal_) hash_string.append(std::to_string(diagonal_->hash()));
    
    hash_string.append(vertical_ ? "1" : "0");
    if (vertical_) hash_string.append(std::to_string(vertical_->hash()));
    
    hash_string.append(horizontal_ ? "1" : "0");
    if (horizontal_) hash_string.append(std::to_string(horizontal_->hash()));

    return hash_string;
}

} // namespace xlnt
