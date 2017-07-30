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
#pragma once

#include <list>
#include <string>
#include <vector>

#include <detail/implementations/conditional_format_impl.hpp>
#include <detail/implementations/format_impl.hpp>
#include <detail/implementations/style_impl.hpp>
#include <xlnt/cell/cell.hpp>
#include <xlnt/styles/conditional_format.hpp>
#include <xlnt/styles/format.hpp>
#include <xlnt/styles/style.hpp>
#include <xlnt/utils/exceptions.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/workbook/worksheet_iterator.hpp>
#include <xlnt/worksheet/worksheet.hpp>

namespace xlnt {
namespace detail {

struct stylesheet
{
    class format create_format(bool default_format)
    {
		format_impls.push_back(format_impl());
		auto &impl = format_impls.back();

		impl.parent = this;
		impl.id = format_impls.size() - 1;
        
        impl.border_id = 0;
        impl.fill_id = 0;
        impl.font_id = 0;
        impl.number_format_id = 0;
        
        impl.references = default_format ? 1 : 0;
        
        return xlnt::format(&impl);
    }

    class xlnt::format format(std::size_t index)
    {
        auto iter = format_impls.begin();
        std::advance(iter, static_cast<std::list<format_impl>::difference_type>(index));

        return xlnt::format(&*iter);
    }

    class style create_style(const std::string &name)
    {
        auto &impl = style_impls.emplace(name, style_impl()).first->second;

		impl.parent = this;
        impl.name = name;

        impl.border_id = 0;
        impl.fill_id = 0;
        impl.font_id = 0;
        impl.number_format_id = 0;

        style_names.push_back(name);
        return xlnt::style(&impl);
    }

	class style create_builtin_style(const std::size_t builtin_id)
	{
		// From Annex G.2
		static const auto *names = new std::unordered_map<std::size_t , std::string>
		{
			{ 0, "Normal" },
			{ 1, "RowLevel_1" },
			{ 2, "ColLevel_1" },
			{ 3, "Comma" },
			{ 4, "Currency" },
			{ 5, "Percent" },
			{ 6, "Comma [0]" },
			{ 7, "Currency [0]" },
			{ 8, "Hyperlink" },
			{ 9, "Followed Hyperlink" },
			{ 10, "Note" },
			{ 11, "Warning Text" },
			{ 15, "Title" },
			{ 16, "Heading 1" },
			{ 17, "Heading 2" },
			{ 18, "Heading 3" },
			{ 19, "Heading 4" },
			{ 20, "Input" },
			{ 21, "Output"},
			{ 22, "Calculation"},
			{ 22, "Calculation" },
			{ 23, "Check Cell" },
			{ 24, "Linked Cell" },
			{ 25, "Total" },
			{ 26, "Good" },
			{ 27, "Bad" },
			{ 28, "Neutral" },
			{ 29, "Accent1" },
			{ 30, "20% - Accent1" },
			{ 31, "40% - Accent1" },
			{ 32, "60% - Accent1" },
			{ 33, "Accent2" },
			{ 34, "20% - Accent2" },
			{ 35, "40% - Accent2" },
			{ 36, "60% - Accent2" },
			{ 37, "Accent3" },
			{ 38, "20% - Accent3" },
			{ 39, "40% - Accent3" },
			{ 40, "60% - Accent3" },
			{ 41, "Accent4" },
			{ 42, "20% - Accent4" },
			{ 43, "40% - Accent4" },
			{ 44, "60% - Accent4" },
			{ 45, "Accent5" },
			{ 46, "20% - Accent5" },
			{ 47, "40% - Accent5" },
			{ 48, "60% - Accent5" },
			{ 49, "Accent6" },
			{ 50, "20% - Accent6" },
			{ 51, "40% - Accent6" },
			{ 52, "60% - Accent6" },
			{ 53, "Explanatory Text" }
		};

		auto new_style = create_style(names->at(builtin_id));
		new_style.d_->builtin_id = builtin_id;

		return new_style;
	}

    class style style(const std::string &name)
	{
        if (!has_style(name)) throw key_not_found();
        return xlnt::style(&style_impls[name]);
	}

	bool has_style(const std::string &name)
	{
		return style_impls.count(name) > 0;
	}

	std::size_t next_custom_number_format_id() const
	{
		std::size_t id = 164;

		for (const auto &nf : number_formats)
		{
			if (nf.id() >= id)
			{
				id = nf.id() + 1;
			}
		}

		return id;
	}
    
    template<typename T, typename C>
    std::size_t find_or_add(C &container, const T &item, bool *added = nullptr)
    {
        if (added != nullptr)
        {
            *added = false;
        }

        std::size_t i = 0;
        
        for (auto iter = container.begin(); iter != container.end(); ++iter)
        {
            if (*iter == item)
            {
                return i;
            }

            ++i;
        }

        if (added != nullptr)
        {
            *added = true;
        }

        container.emplace(container.end(), item);

        return container.size() - 1;
    }
    
    template<typename T>
    std::unordered_map<std::size_t, std::size_t> garbage_collect(
        const std::unordered_map<std::size_t, std::size_t> &reference_counts,
        std::vector<T> &container)
    {
        std::unordered_map<std::size_t, std::size_t> id_map;
        std::size_t unreferenced = 0;
        const auto original_size = container.size();

        for (std::size_t i = 0; i < original_size; ++i)
        {
            id_map[i] = i - unreferenced;

            if (reference_counts.count(i) == 0 || reference_counts.at(i) == 0)
            {
                container.erase(container.begin() + static_cast<typename std::vector<T>::difference_type>(i - unreferenced));
                unreferenced++;
            }
        }

        return id_map;
    }
    
    void garbage_collect()
    {
        if (!garbage_collection_enabled) return;
        
        auto format_iter = format_impls.begin();

        while (format_iter != format_impls.end())
        {
            auto &impl = *format_iter;

            if (impl.references != 0)
            {
                ++format_iter;
            }
            else
            {
                format_iter = format_impls.erase(format_iter);
            }
        }
        
        std::size_t new_id = 0;

        std::unordered_map<std::size_t, std::size_t> alignment_reference_counts;
        std::unordered_map<std::size_t, std::size_t> border_reference_counts;
        std::unordered_map<std::size_t, std::size_t> fill_reference_counts;
        std::unordered_map<std::size_t, std::size_t> font_reference_counts;
        std::unordered_map<std::size_t, std::size_t> number_format_reference_counts;
        std::unordered_map<std::size_t, std::size_t> protection_reference_counts;
        
        fill_reference_counts[0]++;
        fill_reference_counts[1]++;
        
        for (auto &impl : format_impls)
        {
            impl.id = new_id++;
            
            if (impl.alignment_id.is_set())
            {
                alignment_reference_counts[impl.alignment_id.get()]++;
            }
            
            if (impl.border_id.is_set())
            {
                border_reference_counts[impl.border_id.get()]++;
            }
            
            if (impl.fill_id.is_set())
            {
                fill_reference_counts[impl.fill_id.get()]++;
            }
            
            if (impl.font_id.is_set())
            {
                font_reference_counts[impl.font_id.get()]++;
            }
            
            if (impl.number_format_id.is_set())
            {
                number_format_reference_counts[impl.number_format_id.get()]++;
            }
            
            if (impl.protection_id.is_set())
            {
                protection_reference_counts[impl.protection_id.get()]++;
            }
        }
        
        for (auto &name_impl_pair : style_impls)
        {
            auto &impl = name_impl_pair.second;
            
            if (impl.alignment_id.is_set())
            {
                alignment_reference_counts[impl.alignment_id.get()]++;
            }
            
            if (impl.border_id.is_set())
            {
                border_reference_counts[impl.border_id.get()]++;
            }
            
            if (impl.fill_id.is_set())
            {
                fill_reference_counts[impl.fill_id.get()]++;
            }
            
            if (impl.font_id.is_set())
            {
                font_reference_counts[impl.font_id.get()]++;
            }
            
            if (impl.number_format_id.is_set())
            {
                number_format_reference_counts[impl.number_format_id.get()]++;
            }
            
            if (impl.protection_id.is_set())
            {
                protection_reference_counts[impl.protection_id.get()]++;
            }
        }
        
        auto alignment_id_map = garbage_collect(alignment_reference_counts, alignments);
        auto border_id_map = garbage_collect(border_reference_counts, borders);
        auto fill_id_map = garbage_collect(fill_reference_counts, fills);
        auto font_id_map = garbage_collect(font_reference_counts, fonts);
        auto protection_id_map = garbage_collect(protection_reference_counts, protections);

        for (auto &impl : format_impls)
        {
            if (impl.alignment_id.is_set())
            {
                impl.alignment_id = alignment_id_map[impl.alignment_id.get()];
            }
            
            if (impl.border_id.is_set())
            {
                impl.border_id = border_id_map[impl.border_id.get()];
            }
            
            if (impl.fill_id.is_set())
            {
                impl.fill_id = fill_id_map[impl.fill_id.get()];
            }
            
            if (impl.font_id.is_set())
            {
                impl.font_id = font_id_map[impl.font_id.get()];
            }
            
            if (impl.protection_id.is_set())
            {
                impl.protection_id = protection_id_map[impl.protection_id.get()];
            }
        }

        for (auto &name_impl : style_impls)
        {
            auto &impl = name_impl.second;

            if (impl.alignment_id.is_set())
            {
                impl.alignment_id = alignment_id_map[impl.alignment_id.get()];
            }
            
            if (impl.border_id.is_set())
            {
                impl.border_id = border_id_map[impl.border_id.get()];
            }
            
            if (impl.fill_id.is_set())
            {
                impl.fill_id = fill_id_map[impl.fill_id.get()];
            }
            
            if (impl.font_id.is_set())
            {
                impl.font_id = font_id_map[impl.font_id.get()];
            }
            
            if (impl.protection_id.is_set())
            {
                impl.protection_id = protection_id_map[impl.protection_id.get()];
            }
        }
    }

    format_impl *find_or_create(format_impl &pattern)
    {
        auto iter = format_impls.begin();
        bool added = false;
        auto id = find_or_add(format_impls, pattern, &added);
        std::advance(iter, static_cast<std::list<format_impl>::difference_type>(id));
        
        auto &result = *iter;

        if (added)
        {
            result.references = 0;
        }

        result.parent = this;
        result.id = id;
        result.references++;
        
        if (id != pattern.id)
        {
            pattern.references -= pattern.references > 0 ? 1 : 0;
            garbage_collect();
        }

        return &result;
    }

    format_impl *find_or_create_with(format_impl *pattern, const std::string &style_name)
    {
        format_impl new_format = *pattern;
        new_format.style = style_name;

        return find_or_create(new_format);
    }

    format_impl *find_or_create_with(format_impl *pattern, const alignment &new_alignment, bool applied)
    {
        format_impl new_format = *pattern;
        new_format.alignment_id = find_or_add(alignments, new_alignment);
        new_format.alignment_applied = applied;
        
        return find_or_create(new_format);
    }

    format_impl *find_or_create_with(format_impl *pattern, const border &new_border, bool applied)
    {
        format_impl new_format = *pattern;
        new_format.border_id = find_or_add(borders, new_border);
        new_format.border_applied = applied;
        
        return find_or_create(new_format);
    }
    
    format_impl *find_or_create_with(format_impl *pattern, const fill &new_fill, bool applied)
    {
        format_impl new_format = *pattern;
        new_format.fill_id = find_or_add(fills, new_fill);
        new_format.fill_applied = applied;
        
        return find_or_create(new_format);
    }
    
    format_impl *find_or_create_with(format_impl *pattern, const font &new_font, bool applied)
    {
        format_impl new_format = *pattern;
        new_format.font_id = find_or_add(fonts, new_font);
        new_format.font_applied = applied;
        
        return find_or_create(new_format);
    }
    
    format_impl *find_or_create_with(format_impl *pattern, const number_format &new_number_format, bool applied)
    {
        format_impl new_format = *pattern;
        if (new_number_format.id() >= 164)
        {
            find_or_add(number_formats, new_number_format);
        }
        new_format.number_format_id = new_number_format.id();
        new_format.number_format_applied = applied;
        
        return find_or_create(new_format);
    }
    
    format_impl *find_or_create_with(format_impl *pattern, const protection &new_protection, bool applied)
    {
        format_impl new_format = *pattern;
        new_format.protection_id = find_or_add(protections, new_protection);
        new_format.protection_applied = applied;
        
        return find_or_create(new_format);
    }

    std::size_t style_index(const std::string &name) const
    {
        return static_cast<std::size_t>(std::distance(style_names.begin(),
	    std::find(style_names.begin(), style_names.end(), name)));
    }
    
    void clear()
    {
		conditional_format_impls.clear();
        format_impls.clear();
        
        style_impls.clear();
        style_names.clear();
        
        alignments.clear();
        borders.clear();
        fills.clear();
        fonts.clear();
        number_formats.clear();
        protections.clear();
        
        colors.clear();
    }

	conditional_format add_conditional_format_rule(worksheet_impl *ws, const range_reference &ref, const condition &when)
	{
		conditional_format_impls.push_back(conditional_format_impl());

		auto &impl = conditional_format_impls.back();
		impl.when = when;
		impl.parent = this;
		impl.target_sheet = ws;
		impl.target_range = ref;
		impl.differential_format_id = conditional_format_impls.size() - 1;

		return xlnt::conditional_format(&impl);
	}

    workbook *parent;
    
    bool garbage_collection_enabled = true;

	std::list<conditional_format_impl> conditional_format_impls;
    std::list<format_impl> format_impls;
    std::unordered_map<std::string, style_impl> style_impls;
    std::vector<std::string> style_names;

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
