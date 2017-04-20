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

#include <cstdint>
#include <functional>
#include <iostream>
#include <memory>
#include <unordered_map>
#include <vector>

#include <detail/external/include_libstudxml.hpp>
#include <detail/serialization/zstream.hpp>

namespace xlnt {

class color;
class rich_text;
class manifest;
class path;
class relationship;
class variant;
class workbook;
class worksheet;

namespace detail {

class izstream;

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
	/// Read all the files needed from the XLSX archive and initialize all of
	/// the data in the workbook to match.
	/// </summary>
	void populate_workbook();

    /// <summary>
    ///
    /// </summary>
    void read_content_types();

    // Metadata Property Readers

	/// <summary>
	/// Parse the core properties about the current package.
	/// </summary>
	void read_core_properties();

    /// <summary>
    /// Parse the core properties about the current package.
    /// </summary>
    void read_extended_properties();

    /// <summary>
    /// Parse the core properties about the current package.
    /// </summary>
    void read_custom_properties();

	// SpreadsheetML-Specific Package Part Readers

	/// <summary>
	/// Parse the main XML document about the workbook and then all child relationships
	/// of the workbook (e.g. worksheets).
	/// </summary>
	void read_office_document(const std::string &content_type);

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
	/// 
	/// </summary>
	void read_image(const path &part);

    // Common Section Readers

    /// <summary>
    /// Read part from the archive and return a vector of relationships
    /// based on the content of that part.
    /// </summary>
    std::vector<relationship> read_relationships(const path &part);

    /// <summary>
    /// Read a CT_Color from the document currently being parsed.
    /// </summary>
    color read_color();

    /// <summary>
    /// Read a rich text CT_RElt from the document currently being parsed.
    /// </summary>
    rich_text read_rich_text(const xml::qname &parent);

    /// <summary>
    /// Returns true if the givent document type represents an XLSX file.
    /// </summary>
    bool document_type_is_xlsx(const std::string &document_content_type);

    // SAX Parsing Helpers

    /// <summary>
    /// In mixed content XML elements, whitespace before and after is not ignored.
    /// Additionally, if PCDATA spans the boundary of the XML read buffer, it will
    /// be parsed as two separate strings instead of on longer string. This method
    /// will read character data until non-character data is peek()ed from the parser
    /// and returns the combined strings. This should be used when parsing mixed
    /// content to ignore whitespace and whenever character data is expected between
    /// tags.
    /// </summary>
    std::string read_text();

    variant read_variant();

    /// <summary>
    /// Read the part from the archive and parse it as XML. After this is called,
    /// xlsx_consumer::parser() will return a reference to the parser that reads
    /// this part.
    /// </summary>
    void read_part(const std::vector<relationship> &rel_chain);

    /// <summary>
    /// libstudxml will throw an exception if all attributes on an element are not
    /// read with xml::parser::attribute(const std::string &). This should therefore
    /// be called if every remaining attribute should be ignored on an element.
    /// </summary>
    void skip_attributes();

    /// <summary>
    /// Skip attribute name if it exists on the currently parsed element in the XML
    /// parser.
    /// </summary>
    void skip_attribute(const std::string &name);

    /// <summary>
    /// Skip attribute name if it exists on the currently parsed element in the XML
    /// parser.
    /// </summary>
    void skip_attribute(const xml::qname &name);

    /// <summary>
    /// Call skip_attribute on every name in names.
    /// </summary>
    void skip_attributes(const std::vector<xml::qname> &names);

    /// <summary>
    /// Call skip_attribute on every name in names.
    /// </summary>
    void skip_attributes(const std::vector<std::string> &names);

    /// <summary>
    /// Read all content in name until the closing tag is reached.
    /// The closing tag will not be handled after this is called.
    /// </summary>
    void skip_remaining_content(const xml::qname &name);

    /// <summary>
    /// Handles the next event in the XML parser and throws an exception
    /// if it is not the start of an element. Additionally sets the content
    /// type of the element to content.
    /// </summary>
    xml::qname expect_start_element(xml::content content);

    /// <summary>
    /// Handles the next event in the XML parser and throws an exception
    /// if the next element is not named name. Sets the content type of
    /// the element to content.
    /// </summary>
    void expect_start_element(const xml::qname &name, xml::content content);

    /// <summary>
    /// Throws an exception if the next event in the XML parser is not
    /// the end of element called name.
    /// </summary>
    void expect_end_element(const xml::qname &name);

    /// <summary>
    /// Returns true if the top of the parsing stack is called name and
    /// the end of that element hasn't been reached in the XML document.
    /// </summary>
    bool in_element(const xml::qname &name);

    /// <summary>
    /// Handles all start and end namespace events from the current parser
    /// and returns a vector of strings containing the URL for each namespace.
    /// </summary>
    std::vector<std::string> read_namespaces();

    // Properties

	/// <summary>
	/// Convenience method to dereference the pointer to the current parser to avoid
	/// having to use "parser_->" constantly.
	/// </summary>
	xml::parser &parser();

    /// <summary>
    /// Convenience method to access the target workbook's manifest.
    /// </summary>
    class manifest &manifest();

	/// <summary>
	/// The ZIP file containing the files that make up the OOXML package.
	/// </summary>
	std::unique_ptr<izstream> archive_;

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
    
    std::vector<xml::qname> stack_;

    bool preserve_space_ = false;
};

} // namespace detail
} // namespace xlnt
