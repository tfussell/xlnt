#include <detail/custom_value_traits.hpp>

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
#ifdef WIN32
    default: throw xlnt::exception("unhandled");
#endif
    }
}


/// <summary>
/// Returns the string representation of the relationship type.
/// </summary>
std::string to_string(relationship::type t)
{
	switch (t)
	{
	case relationship::type::office_document:
		return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
	case relationship::type::thumbnail:
		return "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail";
	case relationship::type::calculation_chain:
		return "http://purl.oclc.org/ooxml/officeDocument/relationships/calcChain";
	case relationship::type::extended_properties:
		return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
	case relationship::type::core_properties:
		return "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
	case relationship::type::worksheet:
		return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
	case relationship::type::shared_string_table:
		return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
	case relationship::type::styles:
		return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
	case relationship::type::theme:
		return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
	case relationship::type::hyperlink:
		return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
	case relationship::type::chartsheet:
		return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
	case relationship::type::comments:
		return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
	case relationship::type::vml_drawing:
		return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
    case relationship::type::unknown:
        return "unknown";
    case relationship::type::custom_properties:
        return "custom-properties";
    case relationship::type::printer_settings:
        return "printer-settings";
    case relationship::type::connections:
        return "connections";
    case relationship::type::custom_property:
        return "custom-property";
    case relationship::type::custom_xml_mappings:
        return "custom-xml-mappings";
    case relationship::type::dialogsheet:
        return "dialogsheet";
    case relationship::type::drawings:
        return "drawings";
    case relationship::type::external_workbook_references:
        return "external-workbook-references";
    case relationship::type::metadata:
        return "metadata";
    case relationship::type::pivot_table:
        return "pivot-table";
    case relationship::type::pivot_table_cache_definition:
        return "pivot-table-cache-definition";
    case relationship::type::pivot_table_cache_records:
        return "pivot-table-cache-records";
    case relationship::type::query_table:
        return "query-table";
    case relationship::type::shared_workbook_revision_headers:
        return "shared-workbook-revision-headers";
    case relationship::type::shared_workbook:
        return "shared-workbook";
    case relationship::type::revision_log:
        return "revision-log";
    case relationship::type::shared_workbook_user_data:
        return "shared-workbook-user-data";
    case relationship::type::single_cell_table_definitions:
        return "single-cell-table-definitions";
    case relationship::type::table_definition:
        return "table-definition";
    case relationship::type::volatile_dependencies:
        return "volatile-dependencies";
    case relationship::type::image:
        return "image";
#ifdef WIN32
    default:
        throw xlnt::exception("unhandled");
#endif
	}
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
#ifdef WIN32
    default: throw xlnt::exception("unhandled");
#endif
	}
}

std::string to_string(gradient_fill_type fill_type)
{
	return fill_type == gradient_fill_type::linear ? "linear" : "path";
}

std::string to_string(border_style style)
{
	switch (style)
	{
	case border_style::dashdot: return "dashdot";
	case border_style::dashdotdot: return "dashdotdot";
	case border_style::dashed: return "dashed";
	case border_style::dotted: return "dotted";
	case border_style::double_: return "double";
	case border_style::hair: return "hair";
	case border_style::medium: return "medium";
	case border_style::mediumdashdot: return "mediumdashdot";
	case border_style::mediumdashdotdot: return "mediumdashdotdot";
	case border_style::mediumdashed: return "mediumdashed";
	case border_style::slantdashdot: return "slantdashdot";
	case border_style::thick: return "thick";
	case border_style::thin: return "thin";
	case border_style::none: return "none";
#ifdef WIN32
    default: throw xlnt::exception("unhandled");
#endif
	}
}

std::string to_string(vertical_alignment alignment)
{
	switch (alignment)
	{
	case vertical_alignment::bottom: return "bottom";
	case vertical_alignment::center: return "center";
	case vertical_alignment::justify: return "justify";
	case vertical_alignment::top: return "top";
	case vertical_alignment::none: return "none";
#ifdef WIN32
    default: throw xlnt::exception("unhandled");
#endif
	}
}

std::string to_string(horizontal_alignment alignment)
{
	switch (alignment)
	{
	case horizontal_alignment::center: return "center";
	case horizontal_alignment::center_continuous: return "center-continous";
	case horizontal_alignment::general: return "general";
	case horizontal_alignment::justify: return "justify";
	case horizontal_alignment::left: return "left";
	case horizontal_alignment::right: return "right";
	case horizontal_alignment::none: return "none";
#ifdef WIN32
    default: throw xlnt::exception("unhandled");
#endif
	}
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
#ifdef WIN32
    default: throw xlnt::exception("unhandled");
#endif
	}
}

} // namespace detail
} // namespace xlnt
