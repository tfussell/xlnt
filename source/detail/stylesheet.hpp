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
    ~stylesheet() 
	{
	}

    format &create_format()
    {
		format_impls.push_back(format_impl());
		auto &impl = format_impls.back();
		impl.parent = this;
		impl.id = format_impls.size() - 1;
		formats.push_back(format(&impl));
		return formats.back();
    }

	format &get_format(std::size_t index)
	{
		return formats.at(index);
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
			if (nf.get_id() > id)
			{
				id = nf.get_id() + 1;
			}
		}

		return id;
	}

    std::list<format_impl> format_impls;
	std::vector<format> formats;

    std::list<style_impl> style_impls;
	std::vector<style> styles;

	std::vector<alignment> alignments;
    std::vector<border> borders;
    std::vector<fill> fills;
    std::vector<font> fonts;
    std::vector<number_format> number_formats;
	std::vector<protection> protections;
    
    std::vector<color> colors;
};

} // namespace detail
} // namespace xlnt
