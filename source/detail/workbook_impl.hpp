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
#include <xlnt/worksheet/range.hpp>
#include <xlnt/worksheet/range_reference.hpp>
#include <xlnt/worksheet/sheet_view.hpp>

namespace xlnt {
namespace detail {

struct worksheet_impl;

struct workbook_impl
{
	workbook_impl()
		: active_sheet_index_(0),
		guess_types_(false),
		data_only_(false),
		has_theme_(false),
		write_core_properties_(false),
		created_(xlnt::datetime::now()),
		modified_(xlnt::datetime::now()),
		title_("Untitled"),
		base_date_(calendar::windows_1900),
		write_app_properties_(false),
		application_("libxlnt"),
		doc_security_(0),
		scale_crop_(false),
		links_up_to_date_(false),
		shared_doc_(false),
		hyperlinks_changed_(false),
		app_version_("0.9")
	{
	}

    workbook_impl(const workbook_impl &other)
        : active_sheet_index_(other.active_sheet_index_),
          worksheets_(other.worksheets_),
          shared_strings_(other.shared_strings_),
          guess_types_(other.guess_types_),
          data_only_(other.data_only_),
          stylesheet_(other.stylesheet_),
          manifest_(other.manifest_),
		  has_theme_(other.has_theme_),
		  theme_(other.theme_),
		  write_core_properties_(other.write_core_properties_),
		  creator_(other.creator_),
		  last_modified_by_(other.last_modified_by_),
		  created_(other.created_),
		  modified_(other.modified_),
		  title_(other.title_),
		  subject_(other.subject_),
		  description_(other.description_),
		  keywords_(other.keywords_),
		  category_(other.category_),
		  company_(other.company_),
		  base_date_(other.base_date_),
		  write_app_properties_(other.write_app_properties_),
		  application_(other.application_),
		  doc_security_(other.doc_security_),
		  scale_crop_(other.scale_crop_),
		  links_up_to_date_(other.links_up_to_date_),
		  shared_doc_(other.shared_doc_),
		  hyperlinks_changed_(other.hyperlinks_changed_),
		  app_version_(other.app_version_),
		  sheet_title_rel_id_map_(other.sheet_title_rel_id_map_)
    {
    }

    workbook_impl &operator=(const workbook_impl &other)
    {
        active_sheet_index_ = other.active_sheet_index_;
        worksheets_.clear();
        std::copy(other.worksheets_.begin(), other.worksheets_.end(), back_inserter(worksheets_));
        shared_strings_.clear();
        std::copy(other.shared_strings_.begin(), other.shared_strings_.end(), std::back_inserter(shared_strings_));
        guess_types_ = other.guess_types_;
        data_only_ = other.data_only_;
		has_theme_ = other.has_theme_;
		theme_ = other.theme_;
        manifest_ = other.manifest_;

		write_core_properties_ = other.write_core_properties_;
		creator_ = other.creator_;
		last_modified_by_ = other.last_modified_by_;
		created_ = other.created_;
		modified_ = other.modified_;
		title_ = other.title_;
		subject_ = other.subject_;
		description_ = other.description_;
		keywords_ = other.keywords_;
		category_ = other.category_;
		company_ = other.company_;
		base_date_ = other.base_date_;

		write_app_properties_ = other.write_app_properties_;
		application_ = other.application_;
		doc_security_ = other.doc_security_;
		scale_crop_ = other.scale_crop_;
		links_up_to_date_ = other.links_up_to_date_;
		shared_doc_ = other.shared_doc_;
		hyperlinks_changed_ = other.hyperlinks_changed_;
		app_version_ = other.app_version_;

		sheet_title_rel_id_map_ = other.sheet_title_rel_id_map_;

        return *this;
    }

    std::size_t active_sheet_index_;
    std::list<worksheet_impl> worksheets_;
    std::vector<text> shared_strings_;

    bool guess_types_;
    bool data_only_;

    stylesheet stylesheet_;
    
    manifest manifest_;
	bool has_theme_;
    theme theme_;
    std::vector<std::uint8_t> thumbnail_;

	// core properties

	bool write_core_properties_;
	std::string creator_;
	std::string last_modified_by_;
	datetime created_;
	datetime modified_;
	std::string title_;
	std::string subject_;
	std::string description_;
	std::string keywords_;
	std::string category_;
	std::string company_;
	calendar base_date_;

	// application properties

	bool write_app_properties_;
	std::string application_;
	int doc_security_;
	bool scale_crop_;
	bool links_up_to_date_;
	bool shared_doc_;
	bool hyperlinks_changed_;
	std::string app_version_;

	std::unordered_map<std::string, std::string> sheet_title_rel_id_map_;
};

} // namespace detail
} // namespace xlnt
