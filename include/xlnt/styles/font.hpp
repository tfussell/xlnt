// Copyright (c) 2015 Thomas Fussell
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

#include <xlnt/styles/color.hpp>

namespace xlnt {

class font
{
public:
    enum class underline
    {
        none,
        double_,
        double_accounting,
        single,
        single_accounting
    };
    
    void set_bold(bool bold) { bold_ = bold; }
    
    bool operator==(const font &other) const
    {
        return name_ == other.name_;
    }
    
private:
     std::string name_ = "Calibri";
     int size_ = 11;
     bool bold_ = false;
     bool italic_ = false;
     bool superscript_ = false;
     bool subscript_ = false;
     underline underline_ = underline::none;
     bool strikethrough_ = false;
     color color_ = color::black;
};

} // namespace xlnt
