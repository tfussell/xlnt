#pragma once

#include <detail/default_case.hpp>
#include <detail/include_libstudxml.hpp>
#include <xlnt/packaging/relationship.hpp>
#include <xlnt/styles/border.hpp>
#include <xlnt/styles/font.hpp>
#include <xlnt/styles/fill.hpp>
#include <xlnt/styles/horizontal_alignment.hpp>
#include <xlnt/styles/vertical_alignment.hpp>
#include <xlnt/utils/exceptions.hpp>
#include <xlnt/worksheet/pane.hpp>

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

} // namespace detail
} // namespace xlnt

namespace xml {

template <>
struct value_traits<xlnt::font::underline_style>
{
    static xlnt::font::underline_style parse(std::string underline_string, const parser &)
    {
        using xlnt::font;
        
        static auto *underline_styles = new std::vector<font::underline_style>
        {
            font::underline_style::double_,
            font::underline_style::double_accounting,
            font::underline_style::single,
            font::underline_style::single_accounting,
            font::underline_style::none
        };

        for (auto underline_style : *underline_styles)
        {
            if (xlnt::detail::to_string(underline_style) == underline_string)
            {
                return underline_style;
            }
        }
        
        default_case(font::underline_style::none);
    }

    static std::string serialize(xlnt::font::underline_style underline_style, const serializer &)
    {
        return xlnt::detail::to_string(underline_style);
    }
}; // struct value_traits<xlnt::font::underline_style>

template <>
struct value_traits<xlnt::relationship_type>
{
    static xlnt::relationship_type parse(std::string relationship_type_string, const parser &)
    {
        using xlnt::relationship_type;
        
        static auto *relationship_types = new std::vector<relationship_type>
        {
            // Package parts
            relationship_type::core_properties,
            relationship_type::extended_properties,
            relationship_type::custom_properties,
            relationship_type::office_document,
            relationship_type::thumbnail,
            relationship_type::printer_settings,

            // SpreadsheetML parts
            relationship_type::calculation_chain,
            relationship_type::chartsheet,
            relationship_type::comments,
            relationship_type::connections,
            relationship_type::custom_property,
            relationship_type::custom_xml_mappings,
            relationship_type::dialogsheet,
            relationship_type::drawings,
            relationship_type::external_workbook_references,
            relationship_type::pivot_table,
            relationship_type::pivot_table_cache_definition,
            relationship_type::pivot_table_cache_records,
            relationship_type::query_table,
            relationship_type::shared_string_table,
            relationship_type::shared_workbook_revision_headers,
            relationship_type::shared_workbook,
            relationship_type::theme,
            relationship_type::revision_log,
            relationship_type::shared_workbook_user_data,
            relationship_type::single_cell_table_definitions,
            relationship_type::stylesheet,
            relationship_type::table_definition,
            relationship_type::vml_drawing,
            relationship_type::volatile_dependencies,
            relationship_type::worksheet,

            // Worksheet parts
            relationship_type::hyperlink,
            relationship_type::image
        };

        // Core properties relationships can have two different type strings
        if (relationship_type_string == "http://schemas.openxmlformats.org/"
            "officedocument/2006/relationships/metadata/core-properties")
        {
            return relationship_type::core_properties;
        }

        for (auto type : *relationship_types)
        {
            if (xlnt::detail::to_string(type) == relationship_type_string)
            {
                return type;
            }
        }
        
        // ECMA 376-4 Part 1 Section 9.1.7 says consumers shall not fail to load
        // a document with unknown relationships.
        return relationship_type::unknown;
    }

    static std::string serialize(xlnt::relationship_type type, const serializer &)
    {
        return xlnt::detail::to_string(type);
    }
}; // struct value_traits<xlnt::relationship_type>

template <>
struct value_traits<xlnt::pattern_fill_type>
{
    static xlnt::pattern_fill_type parse(std::string fill_type_string, const parser &)
    {
        using xlnt::pattern_fill_type;
        
        static auto *fill_types = new std::vector<pattern_fill_type>
        {
            pattern_fill_type::darkdown,
            pattern_fill_type::darkgray,
            pattern_fill_type::darkgrid,
            pattern_fill_type::darkhorizontal,
            pattern_fill_type::darktrellis,
            pattern_fill_type::darkup,
            pattern_fill_type::darkvertical,
            pattern_fill_type::gray0625,
            pattern_fill_type::gray125,
            pattern_fill_type::lightdown,
            pattern_fill_type::lightgray,
            pattern_fill_type::lightgrid,
            pattern_fill_type::lighthorizontal,
            pattern_fill_type::lighttrellis,
            pattern_fill_type::lightup,
            pattern_fill_type::lightvertical,
            pattern_fill_type::mediumgray,
            pattern_fill_type::none,
            pattern_fill_type::solid
        };

        for (auto fill_type : *fill_types)
        {
            if (xlnt::detail::to_string(fill_type) == fill_type_string)
            {
                return fill_type;
            }
        }
        
        default_case(pattern_fill_type::none);
    }

    static std::string serialize(xlnt::pattern_fill_type fill_type, const serializer &)
    {
        return xlnt::detail::to_string(fill_type);
    }
}; // struct value_traits<xlnt::fill_type>

template <>
struct value_traits<xlnt::gradient_fill_type>
{
    static xlnt::gradient_fill_type parse(std::string fill_type_string, const parser &)
    {
        using xlnt::gradient_fill_type;
        
        static auto *gradient_fill_types = new std::vector<gradient_fill_type>
        {
            gradient_fill_type::linear,
            gradient_fill_type::path
        };

        for (auto fill_type : *gradient_fill_types)
        {
            if (xlnt::detail::to_string(fill_type) == fill_type_string)
            {
                return fill_type;
            }
        }
        
        default_case(gradient_fill_type::linear);
    }

    static std::string serialize(xlnt::gradient_fill_type fill_type, const serializer &)
    {
        return xlnt::detail::to_string(fill_type);
    }
}; // struct value_traits<xlnt::gradient_fill_type>

template <>
struct value_traits<xlnt::border_style>
{
    static xlnt::border_style parse(std::string style_string, const parser &)
    {
        using xlnt::border_style;
        
        static auto *border_styles = new std::vector<border_style>
        {
            border_style::dashdot,
            border_style::dashdotdot,
            border_style::dashed,
            border_style::dotted,
            border_style::double_,
            border_style::hair,
            border_style::medium,
            border_style::mediumdashdot,
            border_style::mediumdashdotdot,
            border_style::mediumdashed,
            border_style::none,
            border_style::slantdashdot,
            border_style::thick,
            border_style::thin
        };

        for (auto style : *border_styles)
        {
            if (xlnt::detail::to_string(style) == style_string)
            {
                return style;
            }
        }
        
        default_case(border_style::none);
    }

    static std::string
    serialize (xlnt::border_style style, const serializer &)
    {
        return xlnt::detail::to_string(style);
    }
}; // struct value_traits<xlnt::border_style>

template <>
struct value_traits<xlnt::vertical_alignment>
{
    static xlnt::vertical_alignment parse(std::string alignment_string, const parser &)
    {
        using xlnt::vertical_alignment;
        
        static auto *vertical_alignments = new std::vector<vertical_alignment>
        {
            vertical_alignment::top,
            vertical_alignment::center,
            vertical_alignment::bottom,
            vertical_alignment::justify,
            vertical_alignment::distributed
        };

        for (auto alignment : *vertical_alignments)
        {
            if (xlnt::detail::to_string(alignment) == alignment_string)
            {
                return alignment;
            }
        }
        
        default_case(vertical_alignment::top);
    }

    static std::string serialize (xlnt::vertical_alignment alignment, const serializer &)
    {
        return xlnt::detail::to_string(alignment);
    }
}; // struct value_traits<xlnt::vertical_alignment>

template <>
struct value_traits<xlnt::horizontal_alignment>
{
    static xlnt::horizontal_alignment parse(std::string alignment_string, const parser &)
    {
        using xlnt::horizontal_alignment;
        
        static auto *horizontal_alignments = new std::vector<horizontal_alignment>
        {
            horizontal_alignment::general,
            horizontal_alignment::left,
            horizontal_alignment::center,
            horizontal_alignment::right,
            horizontal_alignment::fill,
            horizontal_alignment::justify,
            horizontal_alignment::center_continuous,
            horizontal_alignment::distributed
        };

        for (auto alignment : *horizontal_alignments)
        {
            if (xlnt::detail::to_string(alignment) == alignment_string)
            {
                return alignment;
            }
        }
        
        default_case(horizontal_alignment::general);
    }

    static std::string serialize(xlnt::horizontal_alignment alignment, const serializer &)
    {
        return xlnt::detail::to_string(alignment);
    }
}; // struct value_traits<xlnt::horizontal_alignment>

template <>
struct value_traits<xlnt::border_side>
{
    static xlnt::border_side parse(std::string side_string, const parser &)
    {
        using xlnt::border_side;
        
        static auto *border_sides = new std::vector<border_side>
        {
            border_side::top,
            border_side::bottom,
            border_side::start,
            border_side::end,
            border_side::vertical,
            border_side::horizontal,
            border_side::diagonal
        };

        for (auto side : *border_sides)
        {
            if (xlnt::detail::to_string(side) == side_string)
            {
                return side;
            }
        }

        default_case(xlnt::border_side::top);
    }

    static std::string serialize(xlnt::border_side side, const serializer &)
    {
        return xlnt::detail::to_string(side);
    }
}; // struct value_traits<xlnt::border_side>

template <>
struct value_traits<xlnt::target_mode>
{
    static xlnt::target_mode parse(std::string mode_string, const parser &)
    {
        return mode_string == "Internal" ? xlnt::target_mode::internal : xlnt::target_mode::external;
    }

    static std::string serialize(xlnt::target_mode mode, const serializer &)
    {
        return mode == xlnt::target_mode::internal ? "Internal" : "External";
    }
}; // struct value_traits<xlnt::target_mode>

template <>
struct value_traits<xlnt::pane_state>
{
    static xlnt::pane_state parse(std::string string, const parser &)
    {
        if (string == "frozen") return xlnt::pane_state::frozen;
        else if (string == "frozenSplit") return xlnt::pane_state::frozen_split;
        else if (string == "split") return xlnt::pane_state::split;

        default_case(xlnt::pane_state::frozen);
    }

    static std::string serialize(xlnt::pane_state state, const serializer &)
    {
        switch (state)
        {
        case xlnt::pane_state::frozen: return "frozen";
        case xlnt::pane_state::frozen_split: return "frozenSplit";
        case xlnt::pane_state::split: return "split";
        }

        default_case("frozen");
    }
}; // struct value_traits<xlnt::pane_state>

template <>
struct value_traits<xlnt::pane_corner>
{
    static xlnt::pane_corner parse(std::string string, const parser &)
    {
        if (string == "bottomLeft") return xlnt::pane_corner::bottom_left;
        else if (string == "bottomRight") return xlnt::pane_corner::bottom_right;
        else if (string == "topLeft") return xlnt::pane_corner::top_left;
        else if (string == "topRight") return xlnt::pane_corner::top_right;

        default_case(xlnt::pane_corner::top_left);
    }

    static std::string serialize(xlnt::pane_corner corner, const serializer &)
    {
        switch (corner)
        {
        case xlnt::pane_corner::bottom_left: return "bottomLeft";
        case xlnt::pane_corner::bottom_right: return "bottomRight";
        case xlnt::pane_corner::top_left: return "topLeft";
        case xlnt::pane_corner::top_right: return "topRight";
        }

        default_case("topLeft");
    }
}; // struct value_traits<xlnt::pane_corner>

}


