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

#include <string>
#include <unordered_map>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/packaging/relationship.hpp>
#include <xlnt/utils/path.hpp>

namespace xlnt {

enum class content_type
{
	core_properties,
	extended_properties,
	custom_properties,

	calculation_chain,
	chartsheet,
	comments,
	connections,
	custom_property,
	custom_xml_mappings,
	dialogsheet,
	drawings,
	external_workbook_references,
	metadata,
	pivot_table,
	pivot_table_cache_definition,
	pivot_table_cache_records,
	query_table,
	shared_string_table,
	shared_workbook_revision_headers,
	shared_workbook,
	revision_log,
	shared_workbook_user_data,
	single_cell_table_definitions,
	styles,
	table_definition,
	theme,
	volatile_dependencies,
	workbook,
	worksheet,

	unknown
};

/// <summary>
/// The manifest keeps track of all files in the OOXML package and
/// their type and relationships.
/// </summary>
class XLNT_CLASS manifest
{
public:
	/// <summary>
	/// Convenience method for clear_types() and clear_relationships()
	/// </summary>
    void clear();

	/// <summary>
	/// Convenince method for clear_default_types() and clear_override_types()
	/// </summary>
	void clear_types();

	/// <summary>
	/// Unregisters every default content type (i.e. type based on part extension).
	/// </summary>
	void clear_default_types();

	/// <summary>
	/// Unregisters every content type for every part except default types.
	/// </summary>
	void clear_override_types();

	/// <summary>
	/// Convenince method for clear_package_relationships() and clear_part_relationships()
	/// </summary>
	void clear_relationships();

	/// <summary>
	/// Unregisters every relationship with the package root as the source.
	/// </summary>
	void clear_package_relationships();

	/// <summary>
	/// Unregisters every relationship except those with the package root as the source.
	/// </summary>
	void clear_part_relationships();

	/// <summary>
	/// Returns true if the manifest contains a package relationship with the given type.
	/// </summary>
	bool has_package_relationship(relationship::type type) const;

	/// <summary>
	/// Returns true if the manifest contains a package relationship with the given id.
	/// </summary>
	bool has_package_relationship(const std::string &rel_id) const;

	/// <summary>
	/// Returns true if the manifest contains a relationship with the given type with part as the source.
	/// </summary>
	bool has_part_relationship(const path &part, relationship::type type) const;

	/// <summary>
	/// Returns true if the manifest contains a relationship with the given type with part as the source.
	/// </summary>
	bool has_part_relationship(const path &part, const std::string &rel_id) const;

	relationship get_package_relationship(relationship::type type) const;

	relationship get_package_relationship(const std::string &rel_id) const;

	std::vector<relationship> get_package_relationships() const;

	std::vector<relationship> get_package_relationships(relationship::type type) const;

	relationship get_part_relationship(const path &part, relationship::type type) const;

	relationship get_part_relationship(const path &part, const std::string &rel_id) const;

	std::vector<relationship> get_part_relationships(const path &part) const;

	std::vector<relationship> get_part_relationships(const path &part, relationship::type type) const;

	/// <summary>
	/// Given the path to a part, returns the content type of the part as
	/// a content_type enum value or content_type::unknown if it isn't known.
	/// </summary>
	content_type get_content_type(const path &part) const;

	/// <summary>
	/// Given the path to a part, returns the content type of the part as a string.
	/// </summary>
	std::string get_content_type_string(const path &part) const;

	void register_override_type(const path &part, const std::string &content_type);

	/// <summary>
	/// Unregisters the path of the part of the given type and removes all relationships
	/// with the part as a source or target.
	/// </summary>
	void unregister_override_type(const path &part);

	/// <summary>
	/// Returns true if the part at the given path has been registered in this manifest.
	/// </summary>
	bool has_override_type(const path &part) const;

	std::vector<path> get_overriden_parts() const;

	/// <summary>
	/// 
	/// </summary>
	std::string register_package_relationship(relationship::type type, const path &target_uri, target_mode mode);

	/// <summary>
	/// 
	/// </summary>
	std::string register_package_relationship(relationship::type type, const path &target_uri, target_mode mode, const std::string &rel_id);

	/// <summary>
	/// 
	/// </summary>
	std::string register_part_relationship(const path &source_uri, relationship::type type, const path &target_uri, target_mode mode);

	/// <summary>
	/// 
	/// </summary>
	std::string register_part_relationship(const path &source_uri, relationship::type type, const path &target_uri, target_mode mode, const std::string &rel_id);

	/// <summary>
	/// Returns true if a default content type for the extension has been registered.
	/// </summary>
	bool has_default_type(const std::string &extension) const;

	/// <summary>
	/// Returns a vector of all extensions with registered default content types.
	/// </summary>
	std::vector<std::string> get_default_type_extensions() const;

	std::vector<path> get_parts_with_relationships() const;

	/// <summary>
	/// Returns the registered default content type for parts of the given extension.
	/// </summary>
	std::string get_default_type(const std::string &extension) const;

	/// <summary>
	/// Associates the given extension with the given content type.
	/// </summary>
    void register_default_type(const std::string &extension, const std::string &content_type);

	/// <summary>
	/// Unregisters the default content type for the given extension.
	/// </summary>
	void unregister_default_type(const std::string &extension);

private:
	struct part_info
	{
		std::string content_type;
		std::vector<relationship> relationships;
	};

	std::string next_package_relationship_id() const;

	std::string next_relationship_id(const path &part) const;

	std::unordered_map<path, part_info, std::hash<hashable>> part_infos_;

	std::vector<relationship> package_relationships_;

	std::unordered_map<std::string, std::string> extension_content_type_map_;
};

} // namespace xlnt
