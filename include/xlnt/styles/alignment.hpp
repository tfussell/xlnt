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

/// <summary>
/// Alignment options for use in styles.
/// </summary>
class alignment
{
public:
    enum class horizontal_alignment
    {
        general,
        left,
        right,
        center,
        center_continuous,
        justify
    };
    
    enum class vertical_alignment
    {
        bottom,
        top,
        center,
        justify
    };
    
    void set_wrap_text(bool wrap_text)
    {
        wrap_text_ = wrap_text;
    }
    
    bool operator==(const alignment &other) const
    {
        return hash() == other.hash();
    }
    
    std::size_t hash() const { return 0; }
    
private:
    horizontal_alignment horizontal_ = horizontal_alignment::general;
    vertical_alignment vertical_ = vertical_alignment::bottom;
    int text_rotation_ = 0;
    bool wrap_text_ = false;
    bool shrink_to_fit_ = false;
    int indent_ = 0;
};

} // namespace xlnt
