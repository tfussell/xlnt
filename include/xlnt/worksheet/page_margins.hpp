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

namespace xlnt {

class page_margins
{
  public:
    page_margins() : default_(true), top_(0), left_(0), bottom_(0), right_(0), header_(0), footer_(0)
    {
    }

    bool is_default() const
    {
        return default_;
    }
    double get_top() const
    {
        return top_;
    }
    void set_top(double top)
    {
        default_ = false;
        top_ = top;
    }
    double get_left() const
    {
        return left_;
    }
    void set_left(double left)
    {
        default_ = false;
        left_ = left;
    }
    double get_bottom() const
    {
        return bottom_;
    }
    void set_bottom(double bottom)
    {
        default_ = false;
        bottom_ = bottom;
    }
    double get_right() const
    {
        return right_;
    }
    void set_right(double right)
    {
        default_ = false;
        right_ = right;
    }
    double get_header() const
    {
        return header_;
    }
    void set_header(double header)
    {
        default_ = false;
        header_ = header;
    }
    double get_footer() const
    {
        return footer_;
    }
    void set_footer(double footer)
    {
        default_ = false;
        footer_ = footer;
    }

  private:
    bool default_;
    double top_;
    double left_;
    double bottom_;
    double right_;
    double header_;
    double footer_;
};

} // namespace xlnt
