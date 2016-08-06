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
#pragma once

#include <string>
#include <unordered_map>
#include <vector>

#include <xlnt/styles/border.hpp>
#include <xlnt/styles/fill.hpp>
#include <xlnt/styles/font.hpp>
#include <xlnt/styles/format.hpp>
#include <xlnt/styles/number_format.hpp>
#include <xlnt/styles/style.hpp>

namespace {

template<typename T>
std::size_t index(const std::vector<T> &items, const T &to_find)
{
    auto match = std::find(items.begin(), items.end(), to_find);
    
    if (match == items.end())
    {
        throw std::out_of_range("xlnt::stylesheet index lookup");
    }
    
    return std::distance(items.begin(), match);
}

template<typename T>
bool contains(const std::vector<T> &items, const T &to_find)
{
	auto match = std::find(items.begin(), items.end(), to_find);
	return match != items.end();
}

} // namespace

namespace xlnt {
namespace detail {

struct stylesheet
{
    ~stylesheet() {}
    
    std::size_t index(const format &f) 
	{ 
		return ::index(formats, f);
	}
    
    std::size_t index(const std::string &style_name)
    {
        auto match = std::find_if(styles.begin(), styles.end(),
            [&](const style &s) { return s.get_name() == style_name; });
        
        if (match == styles.end())
        {
            throw std::out_of_range("xlnt::stylesheet style index lookup");
        }
        
        return std::distance(styles.begin(), match);
    }
    
    std::size_t add_format(const format &f)
    {
        auto match = std::find(formats.begin(), formats.end(), f);
        
        if (match != formats.end())
        {
            return std::distance(formats.begin(), match);
        }
        
        auto number_format_match = std::find_if(number_formats.begin(), number_formats.end(), [&](const number_format &nf) { return nf.get_format_string() == f.get_number_format().get_format_string(); });
        
        if (number_format_match == number_formats.end() && f.get_number_format().get_id() >= 164)
        {
            number_formats.push_back(f.get_number_format());
        }
        
        formats.push_back(f);
        format_styles.push_back("Normal");
        
		if (!contains(borders, f.get_border()))
        {
            borders.push_back(f.get_border());
        }

		if (!contains(fills, f.get_fill()))
		{
			fills.push_back(f.get_fill());
		}

		if (!contains(fonts, f.get_font()))
		{
			fonts.push_back(f.get_font());
		}
        
        if (f.get_number_format().get_id() >= 164 && !contains(number_formats, f.get_number_format()))
        {
			number_formats.push_back(f.get_number_format());
        }
        
        return formats.size() - 1;
    }
    
    std::size_t add_style(const style &s)
    {
        auto match = std::find(styles.begin(), styles.end(), s);
        
        if (match != styles.end())
        {
            return std::distance(styles.begin(), match);
        }
        
        styles.push_back(s);
        return styles.size() - 1;
    }

    std::vector<format> formats;
    std::vector<std::string> format_styles;
    std::vector<style> styles;

    std::vector<border> borders;
    std::vector<fill> fills;
    std::vector<font> fonts;
    std::vector<number_format> number_formats;
    
    std::vector<color> colors;

    std::unordered_map<std::size_t, std::string> style_name_map;
    
    std::size_t next_custom_format_id = 164;
};

} // namespace detail
} // namespace xlnt
