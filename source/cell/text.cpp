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

#include <xlnt/cell/text.hpp>
#include <xlnt/cell/text_run.hpp>

namespace xlnt {

void text::clear()
{
	runs_.clear();
}

void text::set_plain_string(const std::string &s)
{
	clear();
	add_run(text_run(s));
}

std::string text::get_plain_string() const
{
	std::string plain_string;

	for (const auto &run : runs_)
	{
		plain_string.append(run.get_string());
	}

	return plain_string;
}

std::vector<text_run> text::get_runs() const
{
	return runs_;
}

void text::add_run(const text_run &t)
{
	runs_.push_back(t);
}

bool text::operator==(const text &rhs) const
{
    if (runs_.size() != rhs.runs_.size()) return false;
    
    for (std::size_t i = 0; i < runs_.size(); i++)
    {
        if (runs_[i].get_string() != rhs.runs_[i].get_string()) return false;
        if (runs_[i].has_formatting() != rhs.runs_[i].has_formatting()) return false;
        
        if (runs_[i].has_formatting())
        {
            if (runs_[i].get_color() != rhs.runs_[i].get_color()) return false;
            if (runs_[i].get_family() != rhs.runs_[i].get_family()) return false;
            if (runs_[i].get_font() != rhs.runs_[i].get_font()) return false;
            if (runs_[i].get_scheme() != rhs.runs_[i].get_scheme()) return false;
            if (runs_[i].get_size() != rhs.runs_[i].get_size()) return false;
        }
    }
    
    return true;
}

} // namespace xlnt
