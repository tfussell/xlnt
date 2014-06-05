// Copyright (c) 2014 Thomas Fussell
// Copyright (c) 2010-2014 openpyxl
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

#include "color.hpp"

namespace xlnt {

class fill
{
public:
    enum class type
    {
        none,
        solid,
        gradient_linear,
        gradient_path,
        pattern_darkdown,
        pattern_darkgray,
        pattern_darkgrid,
        pattern_darkhorizontal,
        pattern_darktrellis,
        pattern_darkup,
        pattern_darkvertical,
        pattern_gray0625,
        pattern_gray125,
        pattern_lightdown,
        pattern_lightgray,
        pattern_lightgrid,
        pattern_lighthorizontal,
        pattern_lighttrellis,
        pattern_lightup,
        pattern_lightvertical,
        pattern_mediumgray,
    };
    
    type type_ = type::none;
    int rotation = 0;
    color start_color = color::white;
    color end_color = color::black;
};

} // namespace xlnt
