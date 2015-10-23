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

#include <cstddef>
#include <functional>

#include "side.hpp"
#include "../common/hash_combine.hpp"

namespace xlnt {
    
enum class diagonal_direction
{
    none,
    up,
    down,
    both
};
    
class border
{
public:
    static border default_border();
    
    bool start_assigned;
    side start;
    bool end_assigned;
    side end;
    bool left_assigned;
    side left;
    bool right_assigned;
    side right;
    bool top_assigned;
    side top;
    bool bottom_assigned;
    side bottom;
    bool diagonal_assigned;
    side diagonal;
    bool vertical_assigned;
    side vertical;
    bool horizontal_assigned;
    side horizontal;

    bool outline = false;
    bool diagonal_up = false;
    bool diagonal_down = false;
    
    diagonal_direction diagonal_direction_ = diagonal_direction::none;
    
    bool operator==(const border &other) const
    {
        return hash() == other.hash();
    }
    
    std::size_t hash() const
    {
        std::size_t seed = 0;
        
        if(start_assigned) hash_combine(seed, start.hash());
        if(end_assigned) hash_combine(seed, end.hash());
        if(left_assigned) hash_combine(seed, left.hash());
        if(right_assigned) hash_combine(seed, right.hash());
        if(top_assigned) hash_combine(seed, top.hash());
        if(bottom_assigned) hash_combine(seed, bottom.hash());
        if(diagonal_assigned) hash_combine(seed, diagonal.hash());
        if(vertical_assigned) hash_combine(seed, vertical.hash());
        if(horizontal_assigned) hash_combine(seed, horizontal.hash());

        return seed;
    }
};

} // namespace xlnt
