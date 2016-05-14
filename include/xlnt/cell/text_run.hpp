// Copyright (c) 2016 Thomas Fussell
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

#include <string>

#include <xlnt/xlnt_config.hpp> // for XLNT_CLASS, XLNT_FUNCTION

namespace xlnt {

class XLNT_CLASS text_run 
{
public:
	text_run();
	text_run(const std::string &string);
    
    bool has_formatting() const;

	std::string get_string() const;
	void set_string(const std::string &string);
    
    std::size_t get_size() const;
    std::string get_color() const;
    std::string get_font() const;
    std::size_t get_family() const;
    std::string get_scheme() const;

private:
	std::string string_;
    bool has_formatting_ = false;
    std::size_t size_;
    std::string color_;
    std::string font_;
    std::size_t family_;
    std::string scheme_;
};

} // namespace xlnt
