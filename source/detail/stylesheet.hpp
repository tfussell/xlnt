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

#include <list>
#include <string>
#include <unordered_map>
#include <vector>

#include <detail/format_impl.hpp>
#include <detail/style_impl.hpp>
#include <xlnt/styles/format.hpp>
#include <xlnt/styles/style.hpp>
#include <xlnt/utils/exceptions.hpp>

namespace xlnt {
namespace detail {

struct stylesheet
{
    format &create_format()
    {
	format_impls.push_back(format_impl());
	auto &impl = format_impls.back();
	impl.parent = this;
	impl.id = format_impls.size() - 1;
	formats.push_back(format(&impl));
        auto &format = formats.back();

        if (!alignments.empty())
        {
            format.alignment(alignments.front(), false);
        }
        else
        {
            format.alignment(alignment(), false);
        }
        
        if (!borders.empty())
        {
            format.border(borders.front(), false);
        }
        
        if (!fills.empty())
        {
            format.fill(fills.front(), false);
        }
        
        if (!fonts.empty())
        {
            format.font(fonts.front(), false);
        }
        
        if (!number_formats.empty())
        {
            format.number_format(number_formats.begin()->second, false);
        }
        
        if (!protections.empty())
        {
            format.protection(protections.front(), false);
        }
        
        return format;
    }

    format &get_format(std::size_t index)
    {
        return formats.at(index);
    }

    std::pair<bool, std::size_t> find_format(format pattern)
    {
        std::size_t index = 0;

        for (const format &f : formats)
        {
            if (f == pattern)
            {
                return {true, index};
            }

            ++index;
        }

        return {false, 0};
    }
    
    style &create_style()
    {
		style_impls.push_back(style_impl());
		auto &impl = style_impls.back();
		impl.parent = this;
		styles.push_back(style(&impl));
		return styles.back();
    }

	style &get_style(const std::string &name)
	{
		for (auto &s : styles)
		{
			if (s.name() == name)
			{
				return s;
			}
		}

		throw key_not_found();
	}

	bool has_style(const std::string &name)
	{
		for (auto &s : styles)
		{
			if (s.name() == name)
			{
				return true;
			}
		}

		return false;
	}

	std::size_t next_custom_number_format_id() const
	{
		std::size_t id = 164;

		for (const auto &nf : number_formats)
		{
			if (nf.second.get_id() >= id)
			{
				id = nf.second.get_id() + 1;
			}
		}

		return id;
	}

    std::size_t add_alignment(const alignment &new_alignment)
    {
        auto match = std::find(alignments.begin(), alignments.end(), new_alignment);

        if (match == alignments.end())
        {
            match = alignments.insert(alignments.end(), new_alignment);
        }

        return std::distance(alignments.begin(), match);
    }
    
    std::size_t add_border(const border &new_border)
    {
        std::size_t index = 0;
        
        for (const border &f : borders)
        {
            if (f == new_border)
            {
                return index;
            }
            
            ++index;
        }
        
        borders.push_back(new_border);

        return index;
    }
    
    std::size_t add_fill(const fill &new_fill)
    {
        std::size_t index = 0;
        
        for (const fill &f : fills)
        {
            if (f == new_fill)
            {
                return index;
            }
            
            ++index;
        }
        
        fills.push_back(new_fill);

        return index;
    }
    
    std::size_t add_font(const font &new_font)
    {
        std::size_t index = 0;
        
        for (const font &f : fonts)
        {
            if (f == new_font)
            {
                return index;
            }
            
            ++index;
        }
        
        fonts.push_back(new_font);

        return index;
    }
    
    std::size_t add_protection(const protection &new_protection)
    {
        std::size_t index = 0;
        
        for (const auto &p : protections)
        {
            if (p == new_protection)
            {
                return index;
            }
            
            ++index;
        }
        
        protections.push_back(new_protection);

        return index;
    }
    
    
    void clear()
    {
        format_impls.clear();
        formats.clear();
        
        style_impls.clear();
        styles.clear();
        
        alignments.clear();
        borders.clear();
        fills.clear();
        fonts.clear();
        number_formats.clear();
        protections.clear();
        
        colors.clear();
    }

    std::list<format_impl> format_impls;
	std::vector<format> formats;

    std::list<style_impl> style_impls;
	std::vector<style> styles;

	std::vector<alignment> alignments;
    std::vector<border> borders;
    std::vector<fill> fills;
    std::vector<font> fonts;
    std::unordered_map<std::size_t, number_format> number_formats;
	std::vector<protection> protections;
    
    std::vector<color> colors;
};

} // namespace detail
} // namespace xlnt
