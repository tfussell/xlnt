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
#include <xlnt/cell/cell.hpp>
#include <xlnt/styles/format.hpp>
#include <xlnt/styles/style.hpp>
#include <xlnt/utils/exceptions.hpp>
#include <xlnt/worksheet/worksheet.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/workbook/worksheet_iterator.hpp>

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
        impl.border_id = 0;
        impl.fill_id = 0;
        impl.font_id = 0;
        impl.number_format_id = 0;

		formats.push_back(format(&impl));
        auto &format = formats.back();
        
        return format;
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
			if (nf.get_id() >= id)
			{
				id = nf.get_id() + 1;
			}
		}

		return id;
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

    format_impl *deduplicate(format_impl *f)
    {
        std::unordered_map<format_impl *, std::size_t> reference_counts;

        for (auto ws : *parent)
        {
            auto dimension = ws.calculate_dimension();

            for (auto row = dimension.get_top_left().get_row(); row <= dimension.get_bottom_right().get_row(); row++)
            {
                for (auto column = dimension.get_top_left().get_column(); column <= dimension.get_bottom_right().get_column(); column++)
                {
                    if (ws.has_cell(cell_reference(column, row)))
                    {
                        auto cell = ws.get_cell(cell_reference(column, row));

                        if (cell.has_format())
                        {
                            reference_counts[formats.at(cell.get_format().id()).d_]++;
                        }
                    }
                }
            }
        }

        auto result = f;

        for (auto &impl : format_impls)
        {
            if (reference_counts[&impl] == 0 && &impl != f)
            {
                result = &impl;
                break;
            }
        }

        /*
        if (result != f)
        {
            auto impl_iter = std::find_if(format_impls.begin(), format_impls.end(),
                [=](const format_impl &impl) { return &impl == f; });
            auto index = std::distance(format_impls.begin(), impl_iter);
            formats.erase(formats.begin() + index);
            format_impls.erase(impl_iter);
        }
        */

        return f;
    }

    workbook *parent;

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
