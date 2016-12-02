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
#include <numeric>

#include <xlnt/cell/formatted_text.hpp>
#include <xlnt/cell/text_run.hpp>

namespace xlnt {

void formatted_text::clear()
{
    runs_.clear();
}

void formatted_text::plain_text(const std::string &s)
{
    clear();
    add_run(text_run(s));
}

std::string formatted_text::plain_text() const
{
    return std::accumulate(runs_.begin(), runs_.end(), std::string(),
        [](const std::string &a, const text_run &run) { return a + run.string(); });
}

std::vector<text_run> formatted_text::runs() const
{
    return runs_;
}

void formatted_text::add_run(const text_run &t)
{
    runs_.push_back(t);
}

bool formatted_text::operator==(const formatted_text &rhs) const
{
    if (runs_.size() != rhs.runs_.size()) return false;
    
    for (std::size_t i = 0; i < runs_.size(); i++)
    {
        if (runs_[i].string() != rhs.runs_[i].string()) return false;
        if (runs_[i].has_formatting() != rhs.runs_[i].has_formatting()) return false;
        
        if (runs_[i].has_formatting())
        {
            if (runs_[i].has_color() != rhs.runs_[i].has_color()
                || (runs_[i].has_color()
                && runs_[i].color() != rhs.runs_[i].color()))
            {
                return false;
            }
            
            if (runs_[i].has_family() != rhs.runs_[i].has_family()
                || (runs_[i].has_family()
                && runs_[i].family() != rhs.runs_[i].family()))
            {
                return false;
            }
            
            if (runs_[i].has_font() != rhs.runs_[i].has_font()
                || (runs_[i].has_font()
                && runs_[i].font() != rhs.runs_[i].font()))
            {
                return false;
            }
            
            if (runs_[i].has_scheme() != rhs.runs_[i].has_scheme()
                || (runs_[i].has_scheme()
                && runs_[i].scheme() != rhs.runs_[i].scheme()))
            {
                return false;
            }
            
            if (runs_[i].has_size() != rhs.runs_[i].has_size()
                || (runs_[i].has_size()
                && runs_[i].size() != rhs.runs_[i].size()))
            {
                return false;
            }
        }
    }
    
    return true;
}

} // namespace xlnt
