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

class borders
{
    struct border
    {
        enum class style
        {
            none,
            dashdot,
            dashdotdot,
            dashed,
            dotted,
            double_,
            hair,
            medium,
            mediumdashdot,
            mediumdashdotdot,
            mediumdashed,
            slantdashdot,
            thick,
            thin
        };
        
        style style_ = style::none;
        color color_ = color::black;
    };
    
    enum class diagonal_direction
    {
        none,
        up,
        down,
        both
    };
    
    border left;
    border right;
    border top;
    border bottom;
    border diagonal;
    //    diagonal_direction diagonal_direction = diagonal_direction::none;
    border all_borders;
    border outline;
    border inside;
    border vertical;
    border horizontal;
};

} // namespace xlnt
