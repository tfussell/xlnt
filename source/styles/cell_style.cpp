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
#include <xlnt/styles/cell_style.hpp>
#include <xlnt/styles/font.hpp>
#include <xlnt/styles/number_format.hpp>
#include <xlnt/styles/protection.hpp>
#include <xlnt/utils/hash_combine.hpp>

namespace xlnt {

cell_style::cell_style()
{
}

cell_style::cell_style(const cell_style &other)
    : common_style(other)
{
}

cell_style &cell_style::operator=(const cell_style &other)
{
    parent_ = other.parent_;
    named_style_name_ = other.named_style_name_;
    
    alignment_ = other.alignment_;
    border_ = other.border_;
    fill_ = other.fill_;
    font_ = other.font_;
    number_format_ = other.number_format_;
    protection_ = other.protection_;

    return *this;
}

std::string cell_style::to_hash_string() const
{
    std::string hash_string("style");
    
    hash_string.append(std::to_string(alignment_.apply()));
    hash_string.append(alignment_.apply() ? std::to_string(alignment_.hash()) : " ");

    hash_string.append(std::to_string(border_.apply()));
    hash_string.append(border_.apply() ? std::to_string(border_.hash()) : " ");

    hash_string.append(std::to_string(font_.apply()));
    hash_string.append(font_.apply() ? std::to_string(font_.hash()) : " ");

    hash_string.append(std::to_string(fill_.apply()));
    hash_string.append(fill_.apply() ? std::to_string(fill_.hash()) : " ");

    hash_string.append(std::to_string(number_format_.apply()));
    hash_string.append(number_format_.apply() ? std::to_string(number_format_.hash()) : " ");

    hash_string.append(std::to_string(protection_.apply()));
    hash_string.append(protection_.apply() ? std::to_string(protection_.hash()) : " ");

    return hash_string;
}

} // namespace xlnt
