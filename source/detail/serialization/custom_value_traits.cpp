// Copyright (c) 2016-2017 Thomas Fussell
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

#include <detail/serialization/custom_value_traits.hpp>

namespace xlnt {
namespace detail {

/// <summary>
/// Returns the string representation of the underline style.
/// </summary>
std::string to_string(font::underline_style style)
{
    switch (style)
    {
    case font::underline_style::double_: return "double";
    case font::underline_style::double_accounting: return "doubleAccounting";
    case font::underline_style::single: return "single";
    case font::underline_style::single_accounting: return "singleAccounting";
    case font::underline_style::none: return "none";
    }

    default_case("single");
}

/// <summary>
/// Returns the string representation of the relationship type.
/// </summary>
std::string to_string(relationship_type t)
{
    switch (t)
    {
    case relationship_type::office_document:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    case relationship_type::thumbnail:
        return "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail";
    case relationship_type::calculation_chain:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";
    case relationship_type::extended_properties:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
    case relationship_type::core_properties:
        return "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
    case relationship_type::worksheet:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
    case relationship_type::shared_string_table:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
    case relationship_type::stylesheet:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    case relationship_type::theme:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
    case relationship_type::hyperlink:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    case relationship_type::chartsheet:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
    case relationship_type::comments:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
    case relationship_type::vml_drawing:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
    case relationship_type::custom_properties:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";
    case relationship_type::printer_settings:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings";
    case relationship_type::connections:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections";
    case relationship_type::custom_property:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customProperty";
    case relationship_type::custom_xml_mappings:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlMappings";
    case relationship_type::dialogsheet:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet";
    case relationship_type::drawings:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
    case relationship_type::external_workbook_references:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath";
    case relationship_type::pivot_table:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";
    case relationship_type::pivot_table_cache_definition:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";
    case relationship_type::pivot_table_cache_records:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords";
    case relationship_type::query_table:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable";
    case relationship_type::shared_workbook_revision_headers:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionHeaders";
    case relationship_type::shared_workbook:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedWorkbook";
    case relationship_type::revision_log:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionLog";
    case relationship_type::shared_workbook_user_data:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/usernames";
    case relationship_type::single_cell_table_definitions:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableSingleCells";
    case relationship_type::table_definition:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
    case relationship_type::volatile_dependencies:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/volatileDependencies";
    case relationship_type::image:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
    case relationship_type::unknown:
        return "unknown";
    }

    default_case("unknown");
}

std::string to_string(pattern_fill_type fill_type)
{
    switch (fill_type)
    {
    case pattern_fill_type::darkdown: return "darkdown";
    case pattern_fill_type::darkgray: return "darkgray";
    case pattern_fill_type::darkgrid: return "darkgrid";
    case pattern_fill_type::darkhorizontal: return "darkhorizontal";
    case pattern_fill_type::darktrellis: return "darkhorizontal";
    case pattern_fill_type::darkup: return "darkup";
    case pattern_fill_type::darkvertical: return "darkvertical";
    case pattern_fill_type::gray0625: return "gray0625";
    case pattern_fill_type::gray125: return "gray125";
    case pattern_fill_type::lightdown: return "lightdown";
    case pattern_fill_type::lightgray: return "lightgray";
    case pattern_fill_type::lightgrid: return "lightgrid";
    case pattern_fill_type::lighthorizontal: return "lighthorizontal";
    case pattern_fill_type::lighttrellis: return "lighttrellis";
    case pattern_fill_type::lightup: return "lightup";
    case pattern_fill_type::lightvertical: return "lightvertical";
    case pattern_fill_type::mediumgray: return "mediumgray";
    case pattern_fill_type::solid: return "solid";
    case pattern_fill_type::none: return "none";
    }

    default_case("none");
}

std::string to_string(gradient_fill_type fill_type)
{
    switch (fill_type)
    {
    case gradient_fill_type::linear: return "linear";
    case gradient_fill_type::path: return "path";
    }

    default_case("linear");
}

std::string to_string(border_style style)
{
    switch (style)
    {
    case border_style::dashdot: return "dashDot";
    case border_style::dashdotdot: return "dashDotDot";
    case border_style::dashed: return "dashed";
    case border_style::dotted: return "dotted";
    case border_style::double_: return "double";
    case border_style::hair: return "hair";
    case border_style::medium: return "medium";
    case border_style::mediumdashdot: return "mediumDashDot";
    case border_style::mediumdashdotdot: return "mediumDashDotDot";
    case border_style::mediumdashed: return "mediumDashed";
    case border_style::slantdashdot: return "slantDashDot";
    case border_style::thick: return "thick";
    case border_style::thin: return "thin";
    case border_style::none: return "none";
    }

    default_case("none");
}

std::string to_string(vertical_alignment alignment)
{
    switch (alignment)
    {
    case vertical_alignment::top: return "top";
    case vertical_alignment::center: return "center";
    case vertical_alignment::bottom: return "bottom";
    case vertical_alignment::justify: return "justify";
    case vertical_alignment::distributed: return "distributed";
    }

    default_case("top");
}

std::string to_string(horizontal_alignment alignment)
{
    switch (alignment)
    {
    case horizontal_alignment::general: return "general";
    case horizontal_alignment::left: return "left";
    case horizontal_alignment::center: return "center";
    case horizontal_alignment::right: return "right";
    case horizontal_alignment::fill: return "fill";
    case horizontal_alignment::justify: return "justify";
    case horizontal_alignment::center_continuous: return "centerContinuous";
    case horizontal_alignment::distributed: return "distributed";
    }

    default_case("general");
}

std::string to_string(border_side side)
{
    switch (side)
    {
    case border_side::bottom: return "bottom";
    case border_side::top: return "top";
    case border_side::start: return "left";
    case border_side::end: return "right";
    case border_side::horizontal: return "horizontal";
    case border_side::vertical: return "vertical";
    case border_side::diagonal: return "diagonal";
    }

    default_case("top");
}

std::string to_string(core_property prop)
{
    switch (prop)
    {
    case core_property::category: return "category";
    case core_property::content_status: return "contentStatus";
    case core_property::created: return "created";
    case core_property::creator: return "creator";
    case core_property::description: return "description";
    case core_property::identifier: return "identifier";
    case core_property::keywords: return "keywords";
    case core_property::language: return "language";
    case core_property::last_modified_by: return "lastModifiedBy";
    case core_property::last_printed: return "lastPrinted";
    case core_property::modified: return "modified";
    case core_property::revision: return "revision";
    case core_property::subject: return "subject";
    case core_property::title: return "title";
    case core_property::version: return "version";
    }

    default_case("category");
}

std::string to_string(extended_property prop)
{
    switch (prop)
    {
    case extended_property::application: return "Application";
    case extended_property::app_version: return "AppVersion";
    case extended_property::characters: return "Characters";
    case extended_property::characters_with_spaces: return "CharactersWithSpaces";
    case extended_property::company: return "Company";
    case extended_property::dig_sig: return "DigSig";
    case extended_property::doc_security: return "DocSecurity";
    case extended_property::heading_pairs: return "HeadingPairs";
    case extended_property::hidden_slides: return "HiddenSlides";
    case extended_property::hyperlinks_changed: return "HyperlinksChanged";
    case extended_property::hyperlink_base: return "HyperlinkBase";
    case extended_property::h_links: return "HLinks";
    case extended_property::lines: return "Lines";
    case extended_property::links_up_to_date: return "LinksUpToDate";
    case extended_property::manager: return "Manager";
    case extended_property::m_m_clips: return "MMClips";
    case extended_property::notes: return "Notes";
    case extended_property::pages: return "Pages";
    case extended_property::paragraphs: return "Paragraphs";
    case extended_property::presentation_format: return "PresentationFormat";
    case extended_property::scale_crop: return "ScaleCrop";
    case extended_property::shared_doc: return "SharedDoc";
    case extended_property::slides: return "Slides";
    case extended_property::template_: return "Template";
    case extended_property::titles_of_parts: return "TitlesOfParts";
    case extended_property::total_time: return "TotalTime";
    case extended_property::words: return "Words";
    }

    default_case("Application");
}

std::string to_string(variant::type type)
{
    switch (type)
    {
    case variant::type::boolean: return "bool";
    case variant::type::date: return "date";
    case variant::type::i4: return "i4";
    case variant::type::lpstr: return "lpstr";
    case variant::type::null: return "null";
    case variant::type::vector: return "vector";
    }

    default_case("null");
}

std::string to_string(pane_corner corner)
{
    switch (corner)
    {
    case pane_corner::bottom_left: return "bottomLeft";
    case pane_corner::bottom_right: return "bottomRight";
    case pane_corner::top_left: return "topLeft";
    case pane_corner::top_right: return "topRight";
    }

    default_case("topLeft");
}

std::string to_string(target_mode mode)
{
    switch (mode)
    {
    case target_mode::external: return "External";
    case target_mode::internal: return "Internal";
    }

    default_case("Internal");
}

std::string to_string(pane_state state)
{
    switch (state)
    {
    case pane_state::frozen: return "frozen";
    case pane_state::frozen_split: return "frozenSplit";
    case pane_state::split: return "split";
    }

    default_case("frozen");
}

} // namespace detail
} // namespace xlnt
