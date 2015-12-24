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
#include <xlnt/worksheet/header_footer.hpp>

namespace xlnt {

header &header_footer::get_left_header()
{
    return left_header_;
}
header &header_footer::get_center_header()
{
    return center_header_;
}
header &header_footer::get_right_header()
{
    return right_header_;
}
footer &header_footer::get_left_footer()
{
    return left_footer_;
}
footer &header_footer::get_center_footer()
{
    return center_footer_;
}
footer &header_footer::get_right_footer()
{
    return right_footer_;
}

bool header_footer::is_default_header() const
{
    return left_header_.is_default() && center_header_.is_default() && right_header_.is_default();
}
bool header_footer::is_default_footer() const
{
    return left_footer_.is_default() && center_footer_.is_default() && right_footer_.is_default();
}
bool header_footer::is_default() const
{
    return is_default_header() && is_default_footer();
}

} // namespace xlnt
