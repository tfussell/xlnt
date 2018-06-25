// Copyright (c) 2016-2018 Thomas Fussell
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

#include <algorithm>
#include <string>
#include <unordered_map>

#include <detail/default_case.hpp>
#include <detail/external/include_libstudxml.hpp>
#include <xlnt/packaging/relationship.hpp>
#include <xlnt/styles/alignment.hpp>
#include <xlnt/styles/border.hpp>
#include <xlnt/styles/font.hpp>
#include <xlnt/styles/fill.hpp>
#include <xlnt/utils/exceptions.hpp>
#include <xlnt/utils/variant.hpp>
#include <xlnt/worksheet/pane.hpp>
#include <xlnt/worksheet/page_setup.hpp>
#include <xlnt/workbook/metadata_property.hpp>

namespace xlnt {
namespace detail {

/// <summary>
/// Returns the string representation of the underline style.
/// </summary>
std::string to_string(font::underline_style underline_style);

/// <summary>
/// Returns the string representation of the relationship type.
/// </summary>
std::string to_string(relationship_type t);

std::string to_string(pattern_fill_type fill_type);

std::string to_string(gradient_fill_type fill_type);

std::string to_string(border_style border_style);

std::string to_string(vertical_alignment vertical_alignment);

std::string to_string(horizontal_alignment horizontal_alignment);

std::string to_string(border_side side);

std::string to_string(core_property prop);

std::string to_string(extended_property prop);

std::string to_string(variant::type type);

std::string to_string(pane_corner corner);

std::string to_string(target_mode mode);

std::string to_string(pane_state state);

std::string to_string(orientation state);

template<typename T>
static T from_string(const std::string &string_value);

template<>
font::underline_style from_string(const std::string &string)
{
    if (string == "double") return font::underline_style::double_;
    if (string == "doubleAccounting") return font::underline_style::double_accounting;
    if (string == "single") return font::underline_style::single;
    if (string == "singleAccounting") return font::underline_style::single_accounting;
    if (string == "none") return font::underline_style::none;

    default_case(font::underline_style::none);
}

template<>
relationship_type from_string(const std::string &string)
{
    if (string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument")
        return relationship_type::office_document;
    else if (string == "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail")
        return relationship_type::thumbnail;
    else if (string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain")
        return relationship_type::calculation_chain;
    else if (string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties")
        return relationship_type::extended_properties;
    else if (string == "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
        || string == "http://schemas.openxmlformats.org/officedocument/2006/relationships/metadata/core-properties")
        return relationship_type::core_properties;
    else if (string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")
        return relationship_type::worksheet;
    else if (string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings")
        return relationship_type::shared_string_table;
    else if (string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles")
        return relationship_type::stylesheet;
    else if (string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
        return relationship_type::theme;
    else if (string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink")
        return relationship_type::hyperlink;
    else if (string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet")
        return relationship_type::chartsheet;
    else if (string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments")
        return relationship_type::comments;
    else if (string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing")
        return relationship_type::vml_drawing;
    else if (string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties")
        return relationship_type::custom_properties;
    else if (string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings")
        return relationship_type::printer_settings;
    else if (string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections")
        return relationship_type::connections;
    else if (string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customProperty")
        return relationship_type::custom_property;
    else if (string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlMappings")
        return relationship_type::custom_xml_mappings;
    else if (string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet")
        return relationship_type::dialogsheet;
    else if (string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing")
        return relationship_type::drawings;
    else if (string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath")
        return relationship_type::external_workbook_references;
    else if (string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable")
        return relationship_type::pivot_table;
    else if (string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition")
        return relationship_type::pivot_table_cache_definition;
    else if (string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords")
        return relationship_type::pivot_table_cache_records;
    else if (string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable")
        return relationship_type::query_table;
    else if (string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionHeaders")
        return relationship_type::shared_workbook_revision_headers;
    else if (string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedWorkbook")
        return relationship_type::shared_workbook;
    else if (string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionLog")
        return relationship_type::revision_log;
    else if (string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/usernames")
        return relationship_type::shared_workbook_user_data;
    else if (string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableSingleCells")
        return relationship_type::single_cell_table_definitions;
    else if (string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table")
        return relationship_type::table_definition;
    else if (string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/volatileDependencies")
        return relationship_type::volatile_dependencies;
    else if (string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
        return relationship_type::image;

    // ECMA 376-4 Part 1 Section 9.1.7 says consumers shall not fail to load
    // a document with unknown relationships.
    return relationship_type::unknown;
}

template<>
pattern_fill_type from_string(const std::string &string)
{
#pragma clang diagnostic push
#pragma clang diagnostic ignored "-Wexit-time-destructors"
    static std::unordered_map<std::string, pattern_fill_type> patternFill {
        {"darkdown", pattern_fill_type::darkdown },
        { "darkgray", pattern_fill_type::darkgray },
        { "darkgrid", pattern_fill_type::darkgrid },
        { "darkhorizontal", pattern_fill_type::darkhorizontal },
        { "darktrellis", pattern_fill_type::darktrellis },
        { "darkup", pattern_fill_type::darkup },
        { "darkvertical", pattern_fill_type::darkvertical },
        { "gray0625", pattern_fill_type::gray0625 },
        { "gray125", pattern_fill_type::gray125 },
        { "lightdown", pattern_fill_type::lightdown },
        { "lightgray", pattern_fill_type::lightgray },
        { "lightgrid", pattern_fill_type::lightgrid },
        { "lighthorizontal", pattern_fill_type::lighthorizontal },
        { "lighttrellis", pattern_fill_type::lighttrellis },
        { "lightup", pattern_fill_type::lightup },
        { "lightvertical", pattern_fill_type::lightvertical },
        { "mediumgray", pattern_fill_type::mediumgray },
        { "none", pattern_fill_type::none },
        { "solid", pattern_fill_type::solid }
    };
#pragma clang diagnostic pop
    
    auto toLower = [](std::string str) {
        auto bg{ std::begin (str) };
        auto en{ std::end (str) };
        std::transform(bg, en, bg, 
                       [](char c) {
                        // static cast to avoid int -> char narrowing warning
                           return static_cast<char>(tolower(c)); 
                       });

        return str;
    };

    auto patternLookup = [](const std::string& key) {
        auto entry { patternFill.find (key) };
        if (entry != std::end (patternFill)) {
            return entry->second;
        }
        else {
            // Note: there won't be an error if there is an unsupported pattern
            return pattern_fill_type::none;
        }
    };

    std::string lowerString {toLower (string) };

    return patternLookup (lowerString);
}

template<>
gradient_fill_type from_string(const std::string &string)
{
    if (string == "linear") return gradient_fill_type::linear;
    else if (string == "path") return gradient_fill_type::path;

    default_case(gradient_fill_type::linear);
}

template<>
border_style from_string(const std::string &string)
{
    if (string == "dashDot") return border_style::dashdot;
    else if (string == "dashDotDot") return border_style::dashdotdot;
    else if (string == "dashed") return border_style::dashed;
    else if (string == "dotted") return border_style::dotted;
    else if (string == "double") return border_style::double_;
    else if (string == "hair") return border_style::hair;
    else if (string == "medium") return border_style::medium;
    else if (string == "mediumDashdot") return border_style::mediumdashdot;
    else if (string == "mediumDashDotDot") return border_style::mediumdashdotdot;
    else if (string == "mediumDashed") return border_style::mediumdashed;
    else if (string == "none") return border_style::none;
    else if (string == "slantDashDot") return border_style::slantdashdot;
    else if (string == "thick") return border_style::thick;
    else if (string == "thin") return border_style::thin;

    default_case(border_style::dashdot);
}

template<>
vertical_alignment from_string(const std::string &string)
{
    if (string == "bottom") return vertical_alignment::bottom;
    else if (string == "center") return vertical_alignment::center;
    else if (string == "distributed") return vertical_alignment::distributed;
    else if (string == "justify") return vertical_alignment::justify;
    else if (string == "top") return vertical_alignment::top;

    default_case(vertical_alignment::top);
}

template<>
horizontal_alignment from_string(const std::string &string)
{
    if (string == "center") return horizontal_alignment::center;
    else if (string == "centerContinuous") return horizontal_alignment::center_continuous;
    else if (string == "distributed") return horizontal_alignment::distributed;
    else if (string == "fill") return horizontal_alignment::fill;
    else if (string == "general") return horizontal_alignment::general;
    else if (string == "justify") return horizontal_alignment::justify;
    else if (string == "left") return horizontal_alignment::left;
    else if (string == "right") return horizontal_alignment::right;

    default_case(horizontal_alignment::general);
}
template<>
border_side from_string(const std::string &string)
{
    if (string == "bottom") return border_side::bottom;
    else if (string == "diagonal") return border_side::diagonal;
    else if (string == "right") return border_side::end;
    else if (string == "horizontal") return border_side::horizontal;
    else if (string == "left") return border_side::start;
    else if (string == "top") return border_side::top;
    else if (string == "vertical") return border_side::vertical;

    default_case(border_side::bottom);
}

template<>
core_property from_string(const std::string &string)
{
    if (string == "category") return core_property::category;
    else if (string == "contentStatus") return core_property::content_status;
    else if (string == "created") return core_property::created;
    else if (string == "creator") return core_property::creator;
    else if (string == "description") return core_property::description;
    else if (string == "identifier") return core_property::identifier;
    else if (string == "keywords") return core_property::keywords;
    else if (string == "language") return core_property::language;
    else if (string == "lastModifiedBy") return core_property::last_modified_by;
    else if (string == "lastPrinted") return core_property::last_printed;
    else if (string == "modified") return core_property::modified;
    else if (string == "revision") return core_property::revision;
    else if (string == "subject") return core_property::subject;
    else if (string == "title") return core_property::title;
    else if (string == "version") return core_property::version;

    default_case(core_property::category);
}

template<>
extended_property from_string(const std::string &string)
{
    if (string == "Application") return extended_property::application;
    else if (string == "AppVersion") return extended_property::app_version;
    else if (string == "Characters") return extended_property::characters;
    else if (string == "CharactersWithSpaces") return extended_property::characters_with_spaces;
    else if (string == "Company") return extended_property::company;
    else if (string == "DigSig") return extended_property::dig_sig;
    else if (string == "DocSecurity") return extended_property::doc_security;
    else if (string == "HeadingPairs") return extended_property::heading_pairs;
    else if (string == "HiddenSlides") return extended_property::hidden_slides;
    else if (string == "HyperlinksChanged") return extended_property::hyperlinks_changed;
    else if (string == "HyperlinkBase") return extended_property::hyperlink_base;
    else if (string == "HLinks") return extended_property::h_links;
    else if (string == "Lines") return extended_property::lines;
    else if (string == "LinksUpToDate") return extended_property::links_up_to_date;
    else if (string == "Manager") return extended_property::manager;
    else if (string == "MMClips") return extended_property::m_m_clips;
    else if (string == "Notes") return extended_property::notes;
    else if (string == "Pages") return extended_property::pages;
    else if (string == "Paragraphs") return extended_property::paragraphs;
    else if (string == "PresentationFormat") return extended_property::presentation_format;
    else if (string == "ScaleCrop") return extended_property::scale_crop;
    else if (string == "SharedDoc") return extended_property::shared_doc;
    else if (string == "Slides") return extended_property::slides;
    else if (string == "Template") return extended_property::template_;
    else if (string == "TitlesOfParts") return extended_property::titles_of_parts;
    else if (string == "TotalTime") return extended_property::total_time;
    else if (string == "Words") return extended_property::words;

    default_case(extended_property::application);
}

/*
template<>
variant::type from_string(const std::string &string)
{
    if (string == "bool") return variant::type::boolean;
    else if (string == "date") return variant::type::date;
    else if (string == "i4") return variant::type::i4;
    else if (string == "lpstr") return variant::type::lpstr;
    else if (string == "null") return variant::type::null;
    else if (string == "vector") return variant::type::vector;

    default_case(variant::type::null);
}
*/

template<>
xlnt::pane_state from_string(const std::string &string)
{
    if (string == "frozen") return xlnt::pane_state::frozen;
    else if (string == "frozenSplit") return xlnt::pane_state::frozen_split;
    else if (string == "split") return xlnt::pane_state::split;

    default_case(xlnt::pane_state::frozen);
}

template<>
target_mode from_string(const std::string &string)
{
    if (string == "Internal") return target_mode::internal;
    else if (string == "External") return target_mode::external;

    default_case(target_mode::internal);
}

template<>
pane_corner from_string(const std::string &string)
{
    if (string == "bottomLeft") return pane_corner::bottom_left;
    else if (string == "bottomRight") return pane_corner::bottom_right;
    else if (string == "topLeft") return pane_corner::top_left;
    else if (string == "topRight") return pane_corner::top_right;

    default_case(pane_corner::bottom_left);
}

template <>
orientation from_string(const std::string &string)
{
    if (string == "default") return orientation::default_orientation;
    else if (string == "landscape") return orientation::landscape;
    else if (string == "portrait") return orientation::portrait;
    default_case(orientation::default);
}

} // namespace detail
} // namespace xlnt

namespace xml {

template <>
struct value_traits<xlnt::font::underline_style>
{
    static xlnt::font::underline_style parse(std::string underline_string, const parser &)
    {
        return xlnt::detail::from_string<xlnt::font::underline_style>(underline_string);
    }

    static std::string serialize(xlnt::font::underline_style underline_style, const serializer &)
    {
        return xlnt::detail::to_string(underline_style);
    }
};

template <>
struct value_traits<xlnt::relationship_type>
{
    static xlnt::relationship_type parse(std::string relationship_type_string, const parser &)
    {
        return xlnt::detail::from_string<xlnt::relationship_type>(relationship_type_string);
    }

    static std::string serialize(xlnt::relationship_type type, const serializer &)
    {
        return xlnt::detail::to_string(type);
    }
};

template <>
struct value_traits<xlnt::pattern_fill_type>
{
    static xlnt::pattern_fill_type parse(std::string fill_type_string, const parser &)
    {
        return xlnt::detail::from_string<xlnt::pattern_fill_type>(fill_type_string);
    }

    static std::string serialize(xlnt::pattern_fill_type fill_type, const serializer &)
    {
        return xlnt::detail::to_string(fill_type);
    }
};

template <>
struct value_traits<xlnt::gradient_fill_type>
{
    static xlnt::gradient_fill_type parse(std::string fill_type_string, const parser &)
    {
        return xlnt::detail::from_string<xlnt::gradient_fill_type>(fill_type_string);
    }

    static std::string serialize(xlnt::gradient_fill_type fill_type, const serializer &)
    {
        return xlnt::detail::to_string(fill_type);
    }
};

template <>
struct value_traits<xlnt::border_style>
{
    static xlnt::border_style parse(std::string style_string, const parser &)
    {
        return xlnt::detail::from_string<xlnt::border_style>(style_string);
    }

    static std::string
    serialize (xlnt::border_style style, const serializer &)
    {
        return xlnt::detail::to_string(style);
    }
};

template <>
struct value_traits<xlnt::vertical_alignment>
{
    static xlnt::vertical_alignment parse(std::string alignment_string, const parser &)
    {
        return xlnt::detail::from_string<xlnt::vertical_alignment>(alignment_string);
    }

    static std::string serialize (xlnt::vertical_alignment alignment, const serializer &)
    {
        return xlnt::detail::to_string(alignment);
    }
};

template <>
struct value_traits<xlnt::horizontal_alignment>
{
    static xlnt::horizontal_alignment parse(std::string alignment_string, const parser &)
    {
        return xlnt::detail::from_string<xlnt::horizontal_alignment>(alignment_string);
    }

    static std::string serialize(xlnt::horizontal_alignment alignment, const serializer &)
    {
        return xlnt::detail::to_string(alignment);
    }
};

template <>
struct value_traits<xlnt::border_side>
{
    static xlnt::border_side parse(std::string side_string, const parser &)
    {
        return xlnt::detail::from_string<xlnt::border_side>(side_string);
    }

    static std::string serialize(xlnt::border_side side, const serializer &)
    {
        return xlnt::detail::to_string(side);
    }
};

template <>
struct value_traits<xlnt::target_mode>
{
    static xlnt::target_mode parse(std::string mode_string, const parser &)
    {
        return xlnt::detail::from_string<xlnt::target_mode>(mode_string);
    }

    static std::string serialize(xlnt::target_mode mode, const serializer &)
    {
        return xlnt::detail::to_string(mode);
    }
};

template <>
struct value_traits<xlnt::pane_state>
{
    static xlnt::pane_state parse(std::string string, const parser &)
    {
        return xlnt::detail::from_string<xlnt::pane_state>(string);
    }

    static std::string serialize(xlnt::pane_state state, const serializer &)
    {
        return xlnt::detail::to_string(state);
    }
};

template <>
struct value_traits<xlnt::pane_corner>
{
    static xlnt::pane_corner parse(std::string string, const parser &)
    {
        return xlnt::detail::from_string<xlnt::pane_corner>(string);
    }

    static std::string serialize(xlnt::pane_corner corner, const serializer &)
    {
        return xlnt::detail::to_string(corner);
    }
};

template <>
struct value_traits<xlnt::core_property>
{
    static xlnt::core_property parse(std::string string, const parser &)
    {
        return xlnt::detail::from_string<xlnt::core_property>(string);
    }

    static std::string serialize(xlnt::core_property corner, const serializer &)
    {
        return xlnt::detail::to_string(corner);
    }
};

template <>
struct value_traits<xlnt::extended_property>
{
    static xlnt::extended_property parse(std::string string, const parser &)
    {
        return xlnt::detail::from_string<xlnt::extended_property>(string);
    }

    static std::string serialize(xlnt::extended_property corner, const serializer &)
    {
        return xlnt::detail::to_string(corner);
    }
};

template <>
struct value_traits<xlnt::orientation>
{
    static xlnt::orientation parse(std::string string, const parser &)
    {
        return xlnt::detail::from_string<xlnt::orientation>(string);
    }

    static std::string serialize(xlnt::orientation orientation, const serializer &)
    {
        return xlnt::detail::to_string(orientation);
    }
};

} // namespace xml


