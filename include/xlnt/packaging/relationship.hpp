// Copyright (c) 2014-2021 Thomas Fussell
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
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE
//
// @license: http://www.opensource.org/licenses/mit-license.php
// @author: see AUTHORS file

#pragma once

#include <string>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/packaging/uri.hpp>
#include <xlnt/utils/path.hpp>

namespace xlnt {

/// <summary>
/// Specifies whether the target of a relationship is inside or outside the Package.
/// </summary>
enum class XLNT_API target_mode
{
    /// <summary>
    /// The relationship references a resource that is external to the package.
    /// </summary>
    internal,
    /// <summary>
    /// The relationship references a part that is inside the package.
    /// </summary>
    external
};

/// <summary>
/// All package relationships must be one of these defined types.
/// </summary>
enum class XLNT_API relationship_type
{
    unknown,

    // Package parts
    core_properties,
    extended_properties,
    custom_properties,
    office_document,
    thumbnail,
    printer_settings,

    // SpreadsheetML parts
    calculation_chain,
    chartsheet,
    comments,
    connections,
    custom_property,
    custom_xml_mappings,
    dialogsheet,
    drawings,
    external_workbook_references,
    pivot_table,
    pivot_table_cache_definition,
    pivot_table_cache_records,
    query_table,
    shared_string_table,
    shared_workbook_revision_headers,
    shared_workbook,
    theme,
    revision_log,
    shared_workbook_user_data,
    single_cell_table_definitions,
    stylesheet,
    table_definition,
    vml_drawing,
    volatile_dependencies,
    worksheet,
    vbaproject,

    // Worksheet parts
    hyperlink,
    image
};

/// <summary>
/// Represents an association between a source Package or part, and a target object which can be a part or external
/// resource.
/// </summary>
class XLNT_API relationship
{
public:
    /// <summary>
    /// Constructs a new empty relationship.
    /// </summary>
    relationship();

    /// <summary>
    /// Constructs a new relationship by specifying all of its properties.
    /// </summary>
    relationship(const std::string &id, relationship_type t, const uri &source,
        const uri &target, xlnt::target_mode mode);

    /// <summary>
    /// Returns a string of the form rId# that identifies the relationship.
    /// </summary>
    const std::string &id() const;

    /// <summary>
    /// Returns the type of this relationship.
    /// </summary>
    relationship_type type() const;

    /// <summary>
    /// Returns whether the target of the relationship is internal or external to the package.
    /// </summary>
    xlnt::target_mode target_mode() const;

    /// <summary>
    /// Returns the URI of the package part this relationship points to.
    /// </summary>
    const uri &source() const;

    /// <summary>
    /// Returns the URI of the package part this relationship points to.
    /// </summary>
    const uri &target() const;

    /// <summary>
    /// Returns true if and only if rhs is equal to this relationship.
    /// </summary>
    bool operator==(const relationship &rhs) const;

    /// <summary>
    /// Returns true if and only if rhs is not equal to this relationship.
    /// </summary>
    bool operator!=(const relationship &rhs) const;

private:
    /// <summary>
    /// The id of this relationship in the format "rId#"
    /// </summary>
    std::string id_;

    /// <summary>
    /// The type of this relationship.
    /// </summary>
    relationship_type type_;

    /// <summary>
    /// The URI of the source of this relationship.
    /// </summary>
    uri source_;

    /// <summary>
    /// The URI of the target of this relationshp.
    /// </summary>
    uri target_;

    /// <summary>
    /// Whether the target of this relationship is internal or external.
    /// </summary>
    xlnt::target_mode mode_;
};

} // namespace xlnt
