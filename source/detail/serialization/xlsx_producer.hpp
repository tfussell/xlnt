// Copyright (c) 2014-2017 Thomas Fussell
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
#include <vector>

#include <detail/constants.hpp>
#include <detail/external/include_libstudxml.hpp>

namespace xml {
class serializer;
} // namespace xml

namespace xlnt {

class border;
class cell;
class cell_reference;
class color;
class fill;
class font;
class path;
class relationship;
class streaming_workbook_writer;
class variant;
class workbook;
class worksheet;

namespace detail {

class ozstream;
struct cell_impl;
struct worksheet_impl;

/// <summary>
/// Handles writing a workbook into an XLSX file.
/// </summary>
class xlsx_producer
{
public:
	xlsx_producer(const workbook &target);

    ~xlsx_producer();

	void write(std::ostream &destination);

    void write(std::ostream &destination, const std::string &password);

private:
    friend class xlnt::streaming_workbook_writer;

    void open(std::ostream &destination);

    cell add_cell(const cell_reference &ref);

    worksheet add_worksheet(const std::string &title);

	/// <summary>
	/// Write all files needed to create a valid XLSX file which represents all
	/// data contained in workbook.
	/// </summary>
	void populate_archive(bool streaming);

    void begin_part(const path &part);
    void end_part();

	// Package Parts

	void write_content_types();
    void write_property(const std::string &name, const variant &value, const std::string &ns, bool custom, std::size_t pid);
	void write_core_properties(const relationship &rel);
    void write_extended_properties(const relationship &rel);
    void write_custom_properties(const relationship &rel);
    void write_image(const path &image_path);

	// SpreadsheetML-Specific Package Parts

	void write_workbook(const relationship &rel);

	// Workbook Relationship Target Parts

	void write_connections(const relationship &rel);
	void write_custom_xml_mappings(const relationship &rel);
	void write_external_workbook_references(const relationship &rel);
	void write_pivot_table(const relationship &rel);
	void write_shared_string_table(const relationship &rel);
	void write_shared_workbook_revision_headers(const relationship &rel);
	void write_shared_workbook(const relationship &rel);
	void write_shared_workbook_user_data(const relationship &rel);
	void write_styles(const relationship &rel);
	void write_theme(const relationship &rel);
	void write_volatile_dependencies(const relationship &rel);

	void write_chartsheet(const relationship &rel);
	void write_dialogsheet(const relationship &rel);
	void write_worksheet(const relationship &rel);

	// Sheet Relationship Target Parts

	void write_comments(const relationship &rel, worksheet ws, const std::vector<cell_reference> &cells);
	void write_vml_drawings(const relationship &rel, worksheet ws, const std::vector<cell_reference> &cells);

	// Other Parts

	void write_custom_property();
	void write_unknown_parts();
	void write_unknown_relationships();

	// Helpers

	/// <summary>
	/// Some XLSX producers write booleans as "true" or "false" while others use "1" and "0".
	/// Both are valid, but we can use this method to write either depending on the producer
	/// we're trying to match.
	/// </summary>
	std::string write_bool(bool boolean) const;
    
    void write_relationships(const std::vector<xlnt::relationship> &relationships, const path &part);
    void write_color(const xlnt::color &color);
    void write_border(const xlnt::border &b);
	void write_fill(const xlnt::fill &f);
	void write_font(const xlnt::font &f);
    void write_table_styles();
    void write_colors(const std::vector<xlnt::color> &colors);

    template<typename T>
    void write_element(const std::string &ns, const std::string &name, T value)
    {
        write_start_element(ns, name);
        write_characters(value);
        write_end_element(ns, name);
    }

    void write_start_element(const std::string &name);

    void write_start_element(const std::string &ns, const std::string &name);

    void write_end_element(const std::string &name);

    void write_end_element(const std::string &ns, const std::string &name);

    void write_namespace(const std::string &ns, const std::string &prefix);

    template<typename T>
    void write_attribute(const std::string &name, T value)
    {
        current_part_serializer_->attribute(name, value);
    }

    template<typename T>
    void write_attribute(const xml::qname &name, T value)
    {
        current_part_serializer_->attribute(name, value);
    }

    template<typename T>
    void write_characters(T characters, bool preserve_whitespace = false)
    {
        if (preserve_whitespace)
        {
            write_attribute(xml::qname(constants::ns("xml"), "space"), "preserve");
        }

        current_part_serializer_->characters(characters);
    }

	/// <summary>
	/// A reference to the workbook which is the object of read/write operations.
	/// </summary>
	const workbook &source_;
    
	std::unique_ptr<ozstream> archive_;
    std::unique_ptr<xml::serializer> current_part_serializer_;
    std::unique_ptr<std::streambuf> current_part_streambuf_;
    std::ostream current_part_stream_;

    bool streaming_ = false;

    std::unique_ptr<detail::cell_impl> streaming_cell_;

    detail::cell_impl *current_cell_;

    detail::worksheet_impl *current_worksheet_;
};

} // namespace detail
} // namespace xlnt
