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
#include <memory>
#include <unordered_map>
#include <vector>

#include <detail/include_libstudxml.hpp>
#include <detail/zip.hpp>

namespace xlnt {

class path;
class relationship;
class workbook;
class worksheet;

namespace detail {

class ZipFileReader;

/// <summary>
/// Handles writing a workbook into an XLSX file.
/// </summary>
class xlsx_consumer
{
public:
	xlsx_consumer(workbook &destination);

	void read(std::istream &source);

	void read(std::istream &source, const std::string &password);

private:
	/// <summary>
	/// Ignore all remaining elements at the same depth in the current XML parser.
	/// </summary>
	void skip_remaining_elements();

	/// <summary>
	/// Convenience method to dereference the pointer to the current parser to avoid
	/// having to use "parser_->" constantly.
	/// </summary>
	xml::parser &parser();

	/// <summary>
	/// Read all the files needed from the XLSX archive and initialize all of
	/// the data in the workbook to match.
	/// </summary>
	void populate_workbook();

	// Package Parts

	/// <summary>
	/// Parse content types ([Content_Types].xml) and package relationships (_rels/.rels).
	/// </summary>
	void read_manifest();

	/// <summary>
	/// Parse the core properties about the current package (usually docProps/core.xml).
	/// </summary>
	void read_core_properties();

	/// <summary>
	/// Parse extra application-speicific package properties (usually docProps/app.xml).
	void read_extended_properties();

	/// <summary>
	/// Parse custom file properties. These aren't associated with a particular application
	/// but extensions to OOXML can use this part to hold extra data about the package.
	/// </summary>
	void read_custom_file_properties();

	// SpreadsheetML-Specific Package Parts

	/// <summary>
	/// Parse the main XML document about the workbook and then all child relationships
	/// of the workbook (e.g. worksheets).
	/// </summary>
	void read_workbook();

	// Workbook Relationship Target Parts

	/// <summary>
	/// xl/calcChain.xml
	/// </summary>
	void read_calculation_chain();

	/// <summary>
	/// 
	/// </summary>
	void read_connections();

	/// <summary>
	/// 
	/// </summary>
	void read_custom_property();

	/// <summary>
	/// 
	/// </summary>
	void read_custom_xml_mappings();

	/// <summary>
	/// 
	/// </summary>
	void read_external_workbook_references();

	/// <summary>
	/// 
	/// </summary>
	void read_metadata();

	/// <summary>
	/// 
	/// </summary>
	void read_pivot_table();

	/// <summary>
	/// xl/sharedStrings.xml
	/// </summary>
	void read_shared_string_table();

	/// <summary>
	/// 
	/// </summary>
	void read_shared_workbook_revision_headers();

	/// <summary>
	/// 
	/// </summary>
	void read_shared_workbook();

	/// <summary>
	/// 
	/// </summary>
	void read_shared_workbook_user_data();

	/// <summary>
	/// xl/styles.xml
	/// </summary>
	void read_stylesheet();

	/// <summary>
	/// xl/theme/theme1.xml
	/// </summary>
	void read_theme();

	/// <summary>
	/// 
	/// </summary>
	void read_volatile_dependencies();

	/// <summary>
	/// xl/sheets/*.xml
	/// </summary>
	void read_chartsheet(const std::string &title);

	/// <summary>
	/// xl/sheets/*.xml
	/// </summary>
	void read_dialogsheet(const std::string &title);

	/// <summary>
	/// xl/sheets/*.xml
	/// </summary>
	void read_worksheet(const std::string &title);

	// Sheet Relationship Target Parts

	/// <summary>
	/// 
	/// </summary>
	void read_comments(worksheet ws);
    
	/// <summary>
	/// 
	/// </summary>
	void read_vml_drawings(worksheet ws);

	/// <summary>
	/// 
	/// </summary>
	void read_drawings();

	// Unknown Parts

	/// <summary>
	/// 
	/// </summary>
	void read_unknown_parts();

	/// <summary>
	/// 
	/// </summary>
	void read_unknown_relationships();

	/// <summary>
	/// The ZIP file containing the files that make up the OOXML package.
	/// </summary>
	std::unique_ptr<ZipFileReader> archive_;

	/// <summary>
	/// Map of sheet titles to relationship IDs.
	/// </summary>
	std::unordered_map<std::string, std::size_t> sheet_title_id_map_;

	/// <summary>
	/// Map of sheet titles to indices. Used to ensure sheets are maintained
	/// in the correct order.
	/// </summary>
	std::unordered_map<std::string, std::size_t> sheet_title_index_map_;

	/// <summary>
	/// A reference to the workbook which is being read.
	/// </summary>
	workbook &target_;

	/// <summary>
	/// This pointer is generally set by instantiating an xml::parser in a function
	/// scope and then calling a read_*() method which uses xlsx_consumer::parser() 
	/// to access the object.
	/// </summary>
	xml::parser *parser_;
};

} // namespace detail
} // namespace xlnt
