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
#include <xlnt/styles/protection.hpp>
#include <xlnt/utils/hash_combine.hpp>

namespace xlnt {

protection::protection() : protection(type::unprotected)
{
}

protection::protection(type t) : locked_(t), hidden_(type::unprotected)
{
}

void protection::set_locked(type locked_type)
{
    locked_ = locked_type;
}

void protection::set_hidden(type hidden_type)
{
    hidden_ = hidden_type;
}

std::string protection::to_hash_string() const
{
    std::string hash_string = "protection";
    
    hash_string.append(std::to_string(static_cast<std::size_t>(locked_)));
    hash_string.append(std::to_string(static_cast<std::size_t>(hidden_)));

    return hash_string;
}

} // namespace xlnt
