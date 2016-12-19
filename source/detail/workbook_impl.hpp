// Copyright (c) 2014-2016 Thomas Fussell
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

#include <detail/stylesheet.hpp>
#include <detail/worksheet_impl.hpp>
#include <xlnt/packaging/manifest.hpp>
#include <xlnt/utils/datetime.hpp>
#include <xlnt/workbook/theme.hpp>
#include <xlnt/workbook/workbook_view.hpp>
#include <xlnt/worksheet/range.hpp>
#include <xlnt/worksheet/range_reference.hpp>
#include <xlnt/worksheet/sheet_view.hpp>

namespace xlnt {
namespace detail {

struct worksheet_impl;

struct workbook_impl
{
	workbook_impl() : active_sheet_index_(0)
	{
	}

    workbook_impl(const workbook_impl &other)
        : active_sheet_index_(other.active_sheet_index_),
          worksheets_(other.worksheets_),
          shared_strings_(other.shared_strings_),
          stylesheet_(other.stylesheet_),
          manifest_(other.manifest_),
		  theme_(other.theme_),
		  view_(other.view_),
		  code_name_(other.code_name_),
		  file_version_(other.file_version_)
    {
    }

    workbook_impl &operator=(const workbook_impl &other)
    {
        active_sheet_index_ = other.active_sheet_index_;
        worksheets_.clear();
        std::copy(other.worksheets_.begin(), other.worksheets_.end(), back_inserter(worksheets_));
        shared_strings_.clear();
        std::copy(other.shared_strings_.begin(), other.shared_strings_.end(), std::back_inserter(shared_strings_));
		theme_ = other.theme_;
        manifest_ = other.manifest_;

		sheet_title_rel_id_map_ = other.sheet_title_rel_id_map_;
		view_ = other.view_;
		code_name_ = other.code_name_;
		file_version_ = other.file_version_;

        return *this;
    }

    optional<std::size_t> active_sheet_index_;

    std::list<worksheet_impl> worksheets_;
    std::vector<formatted_text> shared_strings_;

    optional<stylesheet> stylesheet_;

    calendar base_date_;
    optional<std::string> title_;
    
    manifest manifest_;
    optional<theme> theme_;
    std::unordered_map<std::string, std::vector<std::uint8_t>> images_;

    std::unordered_map<std::string, std::string> core_properties_;
    std::unordered_map<std::string, std::string> extended_properties_;
    std::unordered_map<std::string, std::string> custom_properties_;

	std::unordered_map<std::string, std::string> sheet_title_rel_id_map_;

    optional<workbook_view> view_;
    optional<std::string> code_name_;

	struct file_version_t
	{
		std::string app_name;
		std::size_t last_edited;
		std::size_t lowest_edited;
		std::size_t rup_build;
	};
    
    optional<file_version_t> file_version_;
};

} // namespace detail
} // namespace xlnt
