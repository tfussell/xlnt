// Copyright (c) 2014-2021 Thomas Fussell
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
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE
//
// @license: http://www.opensource.org/licenses/mit-license.php
// @author: see AUTHORS file

#pragma once

#include "constants.hpp"

namespace xlnt {
namespace detail {

/// <summary>
/// Clips the maximum number of reserved elements to a certain upper limit.
/// Information like a "count" is often saved in XLSX files and can be used by std::vector::reserve (or other containers)
/// to allocate the memory right away and thus improve performance. However, malicious or broken files
/// might then cause XLNT to allocate extreme amounts of memory. This function clips the number of elements
/// to an upper limit to protect against such issues, but still allow the caller to pre-allocate memory.
/// </summary>
inline size_t clip_reserve_elements(size_t num_elements)
{
    return std::min(num_elements, xlnt::constants::max_elements_for_reserve());
}

} // namespace detail
} // namespace xlnt
