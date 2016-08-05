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

#include <cstdint>
#include <iostream>
#include <vector>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/packaging/zip_file.hpp>

namespace xlnt {

class path;
class relationship;
class workbook;

namespace detail {

/// <summary>
/// Handles writing a workbook into an XLSX file.
/// </summary>
class XLNT_CLASS xlsx_producer
{
public:
	xlsx_producer(const workbook &target);

	void write(const path &destination);

	void write(std::ostream &destination);

	void write(std::vector<std::uint8_t> &destination);

private:
	/// <summary>
	/// Write all files needed to create a valid XLSX file which represents all
	/// data contained in workbook.
	/// </summary>
	void populate_archive(zip_file &archive);

	// Package Parts

	void write_package_relationships();
	void write_content_types();
	void write_app_properties();
	void write_core_properties();
	void write_custom_file_properties();

	// SpreadsheetML-Specific Package Parts

	void write_workbook();
	void write_workbook_relationships();

	// Workbook Relationship Target Parts

	void write_calculation_chain();
	void write_connections();
	void write_custom_property();
	void write_custom_xml_mappings();
	void write_external_workbook_references();
	void write_metadata();
	void write_pivot_table();
	void write_shared_string_table();
	void write_shared_workbook_revision_headers();
	void write_shared_workbook();
	void write_shared_workbook_user_data();
	void write_styles();
	void write_theme();
	void write_volatile_dependencies();

	void write_chartsheet(const relationship &rel);
	void write_dialogsheet(const relationship &rel);
	bool read_worksheet(const relationship &rel);
	void write_worksheet(const relationship &rel);

	// Sheet Relationship Target Parts

	void write_comments();
	void write_drawings();

	// Unknown Parts

	void write_unknown_parts();
	void write_unknown_relationships();

	/// <summary>
	/// A reference to the workbook which is the object of read/write operations.
	/// </summary>
	const workbook &source_;

	/// <summary>
	/// A reference to the archive into which files representing the workbook
	/// will be written.
	/// </summary>
	zip_file destination_;
};

} // namespace detail
} // namespace xlnt
