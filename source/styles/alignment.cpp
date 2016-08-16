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

#include <xlnt/styles/alignment.hpp>

namespace xlnt {

optional<bool> alignment::wrap() const
{
    return wrap_text_;
}

alignment &alignment::wrap(bool wrap_text)
{
	wrap_text_ = wrap_text;
}

optional<bool> alignment::shrink() const
{
	return shrink_to_fit_;
}

alignment &alignment::shrink(bool shrink_to_fit)
{
    shrink_to_fit_ = shrink_to_fit;
}

optional<horizontal_alignment> alignment::horizontal() const
{
    return horizontal_;
}

alignment &alignment::horizontal(horizontal_alignment horizontal)
{
    horizontal_ = horizontal;
}

optional<vertical_alignment> alignment::vertical() const
{
    return vertical_;
}

alignment &alignment::vertical(vertical_alignment vertical)
{
    vertical_ = vertical;
}

std::string alignment::to_hash_string() const
{
    std::string hash_string;

    hash_string.append(wrap_text_ ? "1" : "0");
    hash_string.append(shrink_to_fit_ ? "1" : "0");
    hash_string.append(std::to_string(static_cast<std::size_t>(horizontal_)));
    hash_string.append(std::to_string(static_cast<std::size_t>(vertical_)));
    hash_string.append(std::to_string(text_rotation_));
    hash_string.append(std::to_string(indent_));
    
    return hash_string;
}

} // namespace xlnt
