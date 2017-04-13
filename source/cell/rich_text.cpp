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
#include <numeric>

#include <xlnt/cell/rich_text.hpp>
#include <xlnt/cell/rich_text_run.hpp>

namespace xlnt {

rich_text::rich_text(const std::string &plain_text)
    : rich_text({plain_text, optional<font>()})
{
}

rich_text::rich_text(const std::string &plain_text, const class font &text_font)
    : rich_text({plain_text, optional<font>(text_font)})
{
}

rich_text::rich_text(const rich_text_run &single_run)
{
    add_run(single_run);
}

void rich_text::clear()
{
    runs_.clear();
}

void rich_text::plain_text(const std::string &s)
{
    clear();
    add_run(rich_text_run{s, {}});
}

std::string rich_text::plain_text() const
{
    return std::accumulate(runs_.begin(), runs_.end(), std::string(),
        [](const std::string &a, const rich_text_run &run) { return a + run.first; });
}

std::vector<rich_text_run> rich_text::runs() const
{
    return runs_;
}

void rich_text::add_run(const rich_text_run &t)
{
    runs_.push_back(t);
}

bool rich_text::operator==(const rich_text &rhs) const
{
    if (runs_.size() != rhs.runs_.size()) return false;

    for (std::size_t i = 0; i < runs_.size(); i++)
    {
        if (runs_[i] != rhs.runs_[i]) return false;
    }

    return true;
}

bool rich_text::operator==(const std::string &rhs) const
{
    return *this == rich_text(rhs);
}

bool rich_text::operator!=(const rich_text &rhs) const
{
    return !(*this == rhs);
}

bool rich_text::operator!=(const std::string &rhs) const
{
    return !(*this == rhs);
}

} // namespace xlnt
