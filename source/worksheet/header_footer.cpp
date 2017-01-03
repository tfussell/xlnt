// Copyright (c) 2014-2017 Thomas Fussell
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

bool header_footer::has_header() const
{
    return !odd_headers_.empty() || !even_headers_.empty() || has_first_page_header();
}

bool header_footer::has_footer() const
{
    return !odd_headers_.empty() || !even_headers_.empty() || has_first_page_footer();
}

bool header_footer::align_with_margins() const
{
    return align_with_margins_;
}

header_footer &header_footer::align_with_margins(bool align)
{
    align_with_margins_ = align;
    return *this;
}

bool header_footer::different_odd_even() const
{
    return different_odd_even_;
}

bool header_footer::different_first() const
{
    return !first_headers_.empty() || !first_footers_.empty();
}

bool header_footer::scale_with_doc() const
{
    return scale_with_doc_;
}

header_footer &header_footer::scale_with_doc(bool scale)
{
    scale_with_doc_ = scale;
    return *this;
}

// Normal Header

bool header_footer::has_header(location where) const
{
    return odd_headers_.count(where) > 0 || has_first_page_header(where);
}

void header_footer::clear_header()
{
    odd_headers_.clear();
    even_headers_.clear();
    first_headers_.clear();
}

void header_footer::clear_header(location where)
{
    odd_headers_.erase(where);
    even_headers_.erase(where);
    first_headers_.erase(where);
}

header_footer &header_footer::header(location where, const std::string &text)
{
    return header(where, rich_text(text));
}

header_footer &header_footer::header(location where, const rich_text &text)
{
    odd_headers_[where] = text;
    return *this;
}

rich_text header_footer::header(location where) const
{
    return odd_headers_.at(where);
}

// First Page Header

bool header_footer::has_first_page_header() const
{
    return !first_headers_.empty();
}

bool header_footer::has_first_page_header(location where) const
{
    return first_headers_.count(where) > 0;
}

void header_footer::clear_first_page_header()
{
    first_headers_.clear();
}

void header_footer::clear_first_page_header(location where)
{
    first_headers_.erase(where);
}

header_footer &header_footer::first_page_header(location where, const rich_text &text)
{
    first_headers_[where] = text;
    return *this;
}

rich_text header_footer::first_page_header(location where) const
{
    return first_headers_.at(where);
}

// Odd/Even Header

bool header_footer::has_odd_even_header() const
{
    return different_odd_even() && !odd_headers_.empty();
}

bool header_footer::has_odd_even_header(location where) const
{
    return different_odd_even() && odd_headers_.count(where) > 0;
}

void header_footer::clear_odd_even_header()
{
    odd_headers_.clear();
    even_headers_.clear();
    different_odd_even_ = false;
}

void header_footer::clear_odd_even_header(location where)
{
    odd_headers_.erase(where);
    even_headers_.erase(where);
}

header_footer &header_footer::odd_even_header(location where, const rich_text &odd, const rich_text &even)
{
    odd_headers_[where] = odd;
    even_headers_[where] = even;
    different_odd_even_ = true;

    return *this;
}

rich_text header_footer::odd_header(location where) const
{
    return odd_headers_.at(where);
}

rich_text header_footer::even_header(location where) const
{
    return even_headers_.at(where);
}

// Normal Footer

bool header_footer::has_footer(location where) const
{
    return odd_footers_.count(where) > 0;
}

void header_footer::clear_footer()
{
    odd_footers_.clear();
    even_footers_.clear();
    first_footers_.clear();
}

void header_footer::clear_footer(location where)
{
    odd_footers_.erase(where);
    even_footers_.erase(where);
    first_footers_.erase(where);
}

header_footer &header_footer::footer(location where, const std::string &text)
{
    return footer(where, rich_text(text));
}

header_footer &header_footer::footer(location where, const rich_text &text)
{
    odd_footers_[where] = text;
    return *this;
}

rich_text header_footer::footer(location where) const
{
    return odd_footers_.at(where);
}

// First Page footer

bool header_footer::has_first_page_footer() const
{
    return different_first() && !first_footers_.empty();
}

bool header_footer::has_first_page_footer(location where) const
{
    return different_first() && first_footers_.count(where) > 0;
}

void header_footer::clear_first_page_footer()
{
    first_footers_.clear();
}

void header_footer::clear_first_page_footer(location where)
{
    first_footers_.erase(where);
}

header_footer &header_footer::first_page_footer(location where, const rich_text &text)
{
    first_footers_[where] = text;
    return *this;
}

rich_text header_footer::first_page_footer(location where) const
{
    return first_footers_.at(where);
}

// Odd/Even Footer

bool header_footer::has_odd_even_footer() const
{
    return different_odd_even() && !even_footers_.empty();
}

bool header_footer::has_odd_even_footer(location where) const
{
    return different_odd_even() && even_footers_.count(where) > 0;
}

void header_footer::clear_odd_even_footer()
{
    odd_footers_.clear();
    even_footers_.clear();
}

void header_footer::clear_odd_even_footer(location where)
{
    odd_footers_.erase(where);
    even_footers_.erase(where);
}

header_footer &header_footer::odd_even_footer(location where, const rich_text &odd, const rich_text &even)
{
    odd_footers_[where] = odd;
    even_footers_[where] = even;
    different_odd_even_ = true;

    return *this;
}

rich_text header_footer::odd_footer(location where) const
{
    return odd_footers_.at(where);
}

rich_text header_footer::even_footer(location where) const
{
    return even_footers_.at(where);
}

} // namespace xlnt
