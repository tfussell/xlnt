// Copyright (c) 2014-2016 Thomas Fussell
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
#include <cmath>
#include <numeric> // for std::accumulate
#include <string>

#include <detail/constants.hpp>
#include <detail/custom_value_traits.hpp>
#include <detail/vector_streambuf.hpp>
#include <detail/workbook_impl.hpp>
#include <detail/xlsx_producer.hpp>
#include <detail/zip.hpp>
#include <xlnt/cell/cell.hpp>
#include <xlnt/utils/path.hpp>
#include <xlnt/packaging/manifest.hpp>
#include <xlnt/workbook/const_worksheet_iterator.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/workbook/workbook_view.hpp>
#include <xlnt/worksheet/worksheet.hpp>

using namespace std::string_literals;

namespace {

/// <summary>
/// Some XML elements do unknown things and will be skipped during writing if this is true.
/// </summary>
const bool skip_unknown_elements = true;

/// <summary>
/// Returns true if d is exactly equal to an integer.
/// </summary>
bool is_integral(long double d)
{
	return std::fabs(d - static_cast<long double>(static_cast<long long int>(d))) == 0.L;
}

std::string fill(const std::string &string, std::size_t length = 2)
{
	if (string.size() >= length)
	{
		return string;
	}

	return std::string(length - string.size(), '0') + string;
}

std::string datetime_to_w3cdtf(const xlnt::datetime &dt)
{
	return std::to_string(dt.year) + "-"
        + fill(std::to_string(dt.month)) + "-"
        + fill(std::to_string(dt.day)) + "T"
        + fill(std::to_string(dt.hour)) + ":"
        + fill(std::to_string(dt.minute)) + ":"
        + fill(std::to_string(dt.second)) + "Z";
}

} // namespace

namespace xlnt {
namespace detail {

xlsx_producer::xlsx_producer(const workbook &target) : source_(target)
{
}

void xlsx_producer::write(std::ostream &destination)
{
	ZipFileWriter archive(destination);
    archive_ = &archive;
	populate_archive();
}

// Part Writing Methods

void xlsx_producer::populate_archive()
{
	write_content_types();
    
    const auto root_rels = source_.get_manifest().get_relationships(path("/"));
    write_relationships(root_rels, path("/"));

	for (auto &rel : root_rels)
	{
        // thumbnail is binary content so we don't want to open an xml serializer stream
        if (rel.get_type() == relationship::type::thumbnail)
        {
            write_thumbnail(rel);
            continue;
        }

        begin_part(rel.get_target().get_path());

		switch (rel.get_type())
		{
		case relationship::type::core_properties:
			write_core_properties(rel);
			break;
            
		case relationship::type::extended_properties:
			write_extended_properties(rel);
			break;
            
		case relationship::type::custom_properties:
			write_custom_properties(rel);
			break;
            
		case relationship::type::office_document:
			write_workbook(rel);
			break;
        
        case relationship::type::thumbnail: break;
        case relationship::type::calculation_chain: break;
        case relationship::type::worksheet: break;
        case relationship::type::shared_string_table: break;
        case relationship::type::styles: break;
        case relationship::type::theme: break;
        case relationship::type::hyperlink: break;
        case relationship::type::chartsheet: break;
        case relationship::type::comments: break;
        case relationship::type::vml_drawing: break;
        case relationship::type::unknown: break;
        case relationship::type::printer_settings: break;
        case relationship::type::connections: break;
        case relationship::type::custom_property: break;
        case relationship::type::custom_xml_mappings: break;
        case relationship::type::dialogsheet: break;
        case relationship::type::drawings: break;
        case relationship::type::external_workbook_references: break;
        case relationship::type::metadata: break;
        case relationship::type::pivot_table: break;
        case relationship::type::pivot_table_cache_definition: break;
        case relationship::type::pivot_table_cache_records: break;
        case relationship::type::query_table: break;
        case relationship::type::shared_workbook_revision_headers: break;
        case relationship::type::shared_workbook: break;
        case relationship::type::revision_log: break;
        case relationship::type::shared_workbook_user_data: break;
        case relationship::type::single_cell_table_definitions: break;
        case relationship::type::table_definition: break;
        case relationship::type::volatile_dependencies: break;
        case relationship::type::image: break;
        }
	}

	// Unknown Parts

	void write_unknown_parts();
	void write_unknown_relationships();

    end_part();
}

void xlsx_producer::end_part()
{
    if (current_part_serializer_)
    {
        current_part_serializer_.reset();
    }
    
    archive_->close();
}

void xlsx_producer::begin_part(const path &part)
{
    end_part();
    current_part_serializer_.reset(new xml::serializer(archive_->open(part.string()), part.string()));
}

// Package Parts

void xlsx_producer::write_content_types()
{
    const auto content_types_path = path("[Content_Types].xml");
    begin_part(content_types_path);

    const auto xmlns = "http://schemas.openxmlformats.org/package/2006/content-types"s;

	serializer().start_element(xmlns, "Types");
	serializer().namespace_decl(xmlns, "");

	for (const auto &extension : source_.get_manifest().get_extensions_with_default_types())
	{
		serializer().start_element(xmlns, "Default");
		serializer().attribute("Extension", extension);
		serializer().attribute("ContentType",
            source_.get_manifest().get_default_type(extension));
		serializer().end_element(xmlns, "Default");
	}

	for (const auto &part : source_.get_manifest().get_parts_with_overriden_types())
	{
		serializer().start_element(xmlns, "Override");
		serializer().attribute("PartName", part.resolve(path("/")).string());
		serializer().attribute("ContentType",
            source_.get_manifest().get_override_type(part));
		serializer().end_element(xmlns, "Override");
	}
    
	serializer().end_element(xmlns, "Types");
}

void xlsx_producer::write_extended_properties(const relationship &/*rel*/)
{
    const auto xmlns = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"s;
    const auto xmlns_vt = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"s;
    
    serializer().start_element(xmlns, "Properties");
    serializer().namespace_decl(xmlns, "");
    serializer().namespace_decl(xmlns_vt, "vt");

    serializer().element(xmlns, "Application", source_.get_application());
    serializer().element(xmlns, "DocSecurity", source_.get_doc_security());
    serializer().element(xmlns, "ScaleCrop", source_.get_scale_crop() ? "true" : "false");

    serializer().start_element(xmlns, "HeadingPairs");
	serializer().start_element(xmlns_vt, "vector");
	serializer().attribute("size", "2");
	serializer().attribute("baseType", "variant");
    serializer().start_element(xmlns_vt, "variant");
    serializer().element(xmlns_vt, "lpstr", "Worksheets");
    serializer().end_element(xmlns_vt, "variant");
    serializer().start_element(xmlns_vt, "variant");
    serializer().element(xmlns_vt, "i4", source_.get_sheet_titles().size());
    serializer().end_element(xmlns_vt, "variant");
    serializer().end_element(xmlns_vt, "vector");
    serializer().end_element(xmlns, "HeadingPairs");

    serializer().start_element(xmlns, "TitlesOfParts");
    serializer().start_element(xmlns_vt, "vector");
    serializer().attribute("size", source_.get_sheet_titles().size());
    serializer().attribute("baseType", "lpstr");

	for (auto ws : source_)
	{
        serializer().element(xmlns_vt, "lpstr", ws.get_title());
	}

    serializer().end_element(xmlns_vt, "vector");
    serializer().end_element(xmlns, "TitlesOfParts");

    serializer().start_element(xmlns, "Company");

	if (!source_.get_company().empty())
	{
        serializer().characters(source_.get_company());
	}
    
    serializer().end_element(xmlns, "Company");

	serializer().element(xmlns, "LinksUpToDate", source_.links_up_to_date() ? "true" : "false");
	serializer().element(xmlns, "SharedDoc", source_.is_shared_doc() ? "true" : "false");
	serializer().element(xmlns, "HyperlinksChanged", source_.hyperlinks_changed() ? "true" : "false");
	serializer().element(xmlns, "AppVersion", source_.get_app_version());
    
    serializer().end_element(xmlns, "Properties");
}

void xlsx_producer::write_core_properties(const relationship &/*rel*/)
{
    static const auto &xmlns_cp = constants::get_namespace("core-properties");
    static const auto &xmlns_dc = constants::get_namespace("dc");
    static const auto &xmlns_dcterms = constants::get_namespace("dcterms");
    static const auto &xmlns_dcmitype = constants::get_namespace("dcmitype");
    static const auto &xmlns_xsi = constants::get_namespace("xsi");
    
	serializer().start_element(xmlns_cp, "coreProperties");
    serializer().namespace_decl(xmlns_cp, "cp");
    serializer().namespace_decl(xmlns_dc, "dc");
    serializer().namespace_decl(xmlns_dcterms, "dcterms");
    serializer().namespace_decl(xmlns_dcmitype, "dcmitype");
    serializer().namespace_decl(xmlns_xsi, "xsi");

	serializer().element(xmlns_dc, "creator", source_.get_creator());
	serializer().element(xmlns_cp, "lastModifiedBy", source_.get_last_modified_by());
	serializer().start_element(xmlns_dcterms, "created");
    serializer().attribute(xmlns_xsi, "type", "dcterms:W3CDTF");
    serializer().characters(datetime_to_w3cdtf(source_.get_created()));
    serializer().end_element(xmlns_dcterms, "created");
    serializer().start_element(xmlns_dcterms, "modified");
	serializer().attribute(xmlns_xsi, "type", "dcterms:W3CDTF");
    serializer().characters(datetime_to_w3cdtf(source_.get_modified()));
    serializer().end_element(xmlns_dcterms, "modified");

	if (!source_.get_title().empty())
	{
		serializer().element(xmlns_dc, "title", source_.get_title());
	}

    if (!skip_unknown_elements)
    {
        serializer().element(xmlns_dc, "description", "");
        serializer().element(xmlns_dc, "subject", "");
        serializer().element(xmlns_cp, "keywords", "");
        serializer().element(xmlns_cp, "category", "");
    }

    serializer().end_element(xmlns_cp, "coreProperties");
}

void xlsx_producer::write_custom_properties(const relationship &/*rel*/)
{
	serializer().element("Properties");
}

// Write SpreadsheetML-Specific Package Parts

void xlsx_producer::write_workbook(const relationship &rel)
{
	std::size_t num_visible = 0;
	bool any_defined_names = false;

	for (auto ws : source_)
	{
		if (!ws.has_page_setup() || ws.get_page_setup().get_sheet_state() == sheet_state::visible)
		{
			num_visible++;
		}

		if (ws.has_auto_filter())
		{
			any_defined_names = true;
		}
	}

	if (num_visible == 0)
	{
		throw no_visible_worksheets();
	}

    static const auto &xmlns = constants::get_namespace("workbook");
    static const auto &xmlns_mc = constants::get_namespace("mc");
    static const auto &xmlns_mx = constants::get_namespace("mx");
    static const auto &xmlns_r = constants::get_namespace("r");
    static const auto &xmlns_s = constants::get_namespace("worksheet");
    static const auto &xmlns_x15 = constants::get_namespace("x15");
    static const auto &xmlns_x15ac = constants::get_namespace("x15ac");

	serializer().start_element(xmlns, "workbook");
    serializer().namespace_decl(xmlns, "");
    serializer().namespace_decl(xmlns_r, "r");

	if (source_.x15_enabled())
	{
        serializer().namespace_decl(xmlns_mc, "mc");
        serializer().namespace_decl(xmlns_x15, "x15");
        serializer().attribute(xmlns_mc, "Ignorable", "x15");
	}

	if (source_.has_file_version())
	{
		serializer().start_element(xmlns, "fileVersion");

		serializer().attribute("appName", source_.get_app_name());
		serializer().attribute("lastEdited", source_.get_last_edited());
		serializer().attribute("lowestEdited", source_.get_lowest_edited());
		serializer().attribute("rupBuild", source_.get_rup_build());
        
		serializer().end_element(xmlns, "fileVersion");
	}

	if (source_.has_properties())
	{
		serializer().start_element(xmlns, "workbookPr");

		if (source_.has_code_name())
		{
			serializer().attribute("codeName", source_.get_code_name());
		}

        if (!skip_unknown_elements)
        {
            serializer().attribute("defaultThemeVersion", "124226");
            serializer().attribute("date1904", source_.get_base_date() == calendar::mac_1904 ? "1" : "0");
        }
        
		serializer().end_element(xmlns, "workbookPr");
	}

	if (source_.has_absolute_path())
	{
		serializer().start_element(xmlns_mc, "AlternateContent");
        serializer().namespace_decl(xmlns_mc, "mc");

		serializer().start_element(xmlns_mc, "Choice");
        serializer().attribute("Requires", "x15");

		serializer().start_element(xmlns_x15ac, "absPath");
        serializer().namespace_decl(xmlns_x15ac, "x15ac");
        serializer().attribute("url", source_.get_absolute_path().string());
        
		serializer().end_element(xmlns_x15ac, "absPath");
		serializer().end_element(xmlns_mc, "Choice");
		serializer().end_element(xmlns_mc, "AlternateContent");
	}

	if (source_.has_view())
	{
		serializer().start_element(xmlns, "bookViews");
		serializer().start_element(xmlns, "workbookView");

		const auto &view = source_.get_view();

        if (!skip_unknown_elements)
        {
            serializer().attribute("activeTab", "0");
            serializer().attribute("autoFilterDateGrouping", "1");
            serializer().attribute("firstSheet", "0");
            serializer().attribute("minimized", "0");
            serializer().attribute("showHorizontalScroll", "1");
            serializer().attribute("showSheetTabs", "1");
            serializer().attribute("showVerticalScroll", "1");
            serializer().attribute("tabRatio", "600");
            serializer().attribute("visibility", "visible");
        }

		serializer().attribute("xWindow", view.x_window);
		serializer().attribute("yWindow", view.y_window);
		serializer().attribute("windowWidth", view.window_width);
		serializer().attribute("windowHeight", view.window_height);
		serializer().attribute("tabRatio", view.tab_ratio);
        
		serializer().end_element(xmlns, "workbookView");
		serializer().end_element(xmlns, "bookViews");
	}

	serializer().start_element(xmlns, "sheets");

	if (any_defined_names)
	{
		serializer().element(xmlns, "definedNames", "");
	}

#pragma clang diagnostic push
#pragma clang diagnostic ignored "-Wrange-loop-analysis"
	for (const auto ws : source_)
	{
		auto sheet_rel_id = source_.d_->sheet_title_rel_id_map_[ws.get_title()];
		auto sheet_rel = source_.d_->manifest_.get_relationship(rel.get_target().get_path(), sheet_rel_id);

		serializer().start_element(xmlns, "sheet");
		serializer().attribute("name", ws.get_title());
		serializer().attribute("sheetId", ws.get_id());

		if (ws.has_page_setup() && ws.get_sheet_state() == xlnt::sheet_state::hidden)
		{
            serializer().attribute("state", "hidden");
		}

        serializer().attribute(xmlns_r, "id", sheet_rel_id);

		if (ws.has_auto_filter())
		{
			serializer().start_element(xmlns, "definedName");
            
			serializer().attribute("name", "_xlnm._FilterDatabase");
			serializer().attribute("hidden", write_bool(true));
			serializer().attribute("localSheetId", "0");
			serializer().characters("'" + ws.get_title() + "'!"
                + range_reference::make_absolute(ws.get_auto_filter()).to_string());
            
			serializer().end_element(xmlns, "definedName");
		}
        
        serializer().end_element(xmlns, "sheet");
	}
#pragma clang diagnostic pop
    
    serializer().end_element(xmlns, "sheets");

	if (source_.has_calculation_properties())
	{
		serializer().start_element(xmlns, "calcPr");
        serializer().attribute("calcId", 150000);
        
        if (!skip_unknown_elements)
        {
            serializer().attribute("calcMode", "auto");
            serializer().attribute("fullCalcOnLoad", "1");
        }
        
        serializer().attribute("concurrentCalc", "0");
        serializer().end_element(xmlns, "calcPr");
	}

	if (!source_.get_named_ranges().empty())
	{
        serializer().start_element(xmlns, "definedNames");

		for (auto &named_range : source_.get_named_ranges())
		{
            serializer().start_element(xmlns_s, "definedName");
            serializer().namespace_decl(xmlns_s, "s");
            serializer().attribute("name", named_range.get_name());
			const auto &target = named_range.get_targets().front();
			serializer().characters("'" + target.first.get_title()
                + "\'!" + target.second.to_string());
            serializer().end_element(xmlns_s, "definedName");
		}
        
        serializer().end_element(xmlns, "definedNames");
	}

	if (source_.has_arch_id())
	{
        serializer().start_element(xmlns, "extLst");
        serializer().start_element(xmlns, "ext");
        serializer().namespace_decl(xmlns_mx, "mx");
        serializer().attribute("uri", "{7523E5D3-25F3-A5E0-1632-64F254C22452}");
        serializer().start_element(xmlns_mx, "ArchID");
        serializer().attribute("Flags", 2);
        serializer().end_element(xmlns_mx, "ArchID");
        serializer().end_element(xmlns, "ext");
        serializer().end_element(xmlns, "extLst");
	}

    serializer().end_element(xmlns, "workbook");
    
    auto workbook_rels = source_.get_manifest().get_relationships(rel.get_target().get_path());
    write_relationships(workbook_rels, rel.get_target().get_path());
    
    for (const auto &child_rel : workbook_rels)
    {
		path archive_path(child_rel.get_source().get_path().parent().append(child_rel.get_target().get_path()));
        begin_part(archive_path);
        
        switch (child_rel.get_type())
        {
        case relationship::type::calculation_chain:
			write_calculation_chain(child_rel);
			break;
            
		case relationship::type::chartsheet:
			write_chartsheet(child_rel);
			break;
            
		case relationship::type::connections:
			write_connections(child_rel);
			break;
            
		case relationship::type::custom_xml_mappings:
			write_custom_xml_mappings(child_rel);
			break;
            
		case relationship::type::dialogsheet:
			write_dialogsheet(child_rel);
			break;
            
		case relationship::type::external_workbook_references:
			write_external_workbook_references(child_rel);
			break;
            
		case relationship::type::metadata:
			write_metadata(child_rel);
			break;
            
		case relationship::type::pivot_table:
			write_pivot_table(child_rel);
			break;

		case relationship::type::shared_string_table:
			write_shared_string_table(child_rel);
			break;

		case relationship::type::shared_workbook_revision_headers:
			write_shared_workbook_revision_headers(child_rel);
			break;
            
		case relationship::type::styles:
			write_styles(child_rel);
			break;
            
		case relationship::type::theme:
			write_theme(child_rel);
			break;
            
		case relationship::type::volatile_dependencies:
			write_volatile_dependencies(child_rel);
			break;
            
		case relationship::type::worksheet:
			write_worksheet(child_rel);
			break;
            
        case relationship::type::office_document: break;
        case relationship::type::thumbnail: break;
        case relationship::type::extended_properties: break;
        case relationship::type::core_properties: break;
        case relationship::type::hyperlink: break;
        case relationship::type::comments: break;
        case relationship::type::vml_drawing: break;
        case relationship::type::unknown: break;
        case relationship::type::custom_properties: break;
        case relationship::type::printer_settings: break;
        case relationship::type::custom_property: break;
        case relationship::type::drawings: break;
        case relationship::type::pivot_table_cache_definition: break;
        case relationship::type::pivot_table_cache_records: break;
        case relationship::type::query_table: break;
        case relationship::type::shared_workbook: break;
        case relationship::type::revision_log: break;
        case relationship::type::shared_workbook_user_data: break;
        case relationship::type::single_cell_table_definitions: break;
        case relationship::type::table_definition: break;
        case relationship::type::image: break;
		}
    }
}

// Write Workbook Relationship Target Parts

void xlsx_producer::write_calculation_chain(const relationship &/*rel*/)
{
	serializer().start_element("calcChain");
}

void xlsx_producer::write_chartsheet(const relationship &/*rel*/)
{
	serializer().start_element("chartsheet");
}

void xlsx_producer::write_connections(const relationship &/*rel*/)
{
	serializer().start_element("connections");
}

void xlsx_producer::write_custom_xml_mappings(const relationship &/*rel*/)
{
	serializer().start_element("MapInfo");
}

void xlsx_producer::write_dialogsheet(const relationship &/*rel*/)
{
	serializer().start_element("dialogsheet");
}

void xlsx_producer::write_external_workbook_references(const relationship &/*rel*/)
{
	serializer().start_element("externalLink");
}

void xlsx_producer::write_metadata(const relationship &/*rel*/)
{
	serializer().start_element("metadata");
}

void xlsx_producer::write_pivot_table(const relationship &/*rel*/)
{
	serializer().start_element("pivotTableDefinition");
}

void xlsx_producer::write_shared_string_table(const relationship &/*rel*/)
{
    static const auto &xmlns = constants::get_namespace("worksheet");

	serializer().start_element(xmlns, "sst");
    serializer().namespace_decl(xmlns, "");

    // todo: is there a more elegant way to get this number?
    std::size_t string_count = 0;

#pragma clang diagnostic push
#pragma clang diagnostic ignored "-Wrange-loop-analysis"
    for (const auto ws : source_)
    {
        auto dimension = ws.calculate_dimension();

        for (xlnt::row_t row = dimension.get_top_left().get_row();
            row <= dimension.get_bottom_right().get_row(); ++row)
        {
            for (xlnt::column_t column = dimension.get_top_left().get_column();
                column <= dimension.get_bottom_right().get_column(); ++column)
            {
                if (ws.has_cell(xlnt::cell_reference(column, row)))
                {
                    string_count += (ws.get_cell(xlnt::cell_reference(column, row))
                        .get_data_type() == cell::type::string) ? 1 : 0;
                }
            }
        }
    }
#pragma clang diagnostic pop

	serializer().attribute("count", string_count);
	serializer().attribute("uniqueCount", source_.get_shared_strings().size());

	for (const auto &string : source_.get_shared_strings())
	{
		if (string.runs().size() == 1 && !string.runs().at(0).has_formatting())
		{
            serializer().start_element(xmlns, "si");
            serializer().element(xmlns, "t", string.plain_text());
            serializer().end_element(xmlns, "si");
            
            continue;
		}

        serializer().start_element(xmlns, "si");

        for (const auto &run : string.runs())
        {
            serializer().start_element(xmlns, "r");

            if (run.has_formatting())
            {
                serializer().start_element(xmlns, "rPr");

                if (run.has_size())
                {
                    serializer().start_element(xmlns, "sz");
                    serializer().attribute("val", run.get_size());
                    serializer().end_element(xmlns, "sz");
                }

                if (run.has_color())
                {
                    serializer().start_element(xmlns, "color");
                    write_color(run.get_color());
                    serializer().end_element(xmlns, "color");
                }

                if (run.has_font())
                {
                    serializer().start_element(xmlns, "rFont");
                    serializer().attribute("val", run.get_font());
                    serializer().end_element(xmlns, "rFont");
                }

                if (run.has_family())
                {
                    serializer().start_element(xmlns, "family");
                    serializer().attribute("val", run.get_family());
                    serializer().end_element(xmlns, "family");
                }

                if (run.has_scheme())
                {
                    serializer().start_element(xmlns, "scheme");
                    serializer().attribute("val", run.get_scheme());
                    serializer().end_element(xmlns, "scheme");
                }

				if (run.bold_set())
				{
					serializer().start_element(xmlns, "b");
					serializer().end_element(xmlns, "b");
				}
                
                serializer().end_element(xmlns, "rPr");
            }

            serializer().element(xmlns, "t", run.get_string());
            serializer().end_element(xmlns, "r");
        }
        
        serializer().end_element(xmlns, "si");
	}

    serializer().end_element(xmlns, "sst");
}

void xlsx_producer::write_shared_workbook_revision_headers(const relationship &/*rel*/)
{
	serializer().start_element("headers");
}

void xlsx_producer::write_shared_workbook(const relationship &/*rel*/)
{
	serializer().start_element("revisions");
}

void xlsx_producer::write_shared_workbook_user_data(const relationship &/*rel*/)
{
	serializer().start_element("users");
}

void xlsx_producer::write_styles(const relationship &/*rel*/)
{
    static const auto &xmlns = constants::get_namespace("worksheet");
    static const auto &xmlns_mc = constants::get_namespace("mc");
    static const auto &xmlns_x14 = constants::get_namespace("x14");
    static const auto &xmlns_x14ac = constants::get_namespace("x14ac");
    
	serializer().start_element(xmlns, "styleSheet");
	serializer().namespace_decl(xmlns, "");

	if (source_.x15_enabled())
	{
        serializer().namespace_decl(xmlns_mc, "mc");
    	serializer().namespace_decl(xmlns_x14ac, "x14ac");
        serializer().attribute(xmlns_mc, "Ignorable", "x14ac");
	}

	const auto &stylesheet = source_.impl().stylesheet_;

	// Number Formats

	if (!stylesheet.number_formats.empty())
	{
        const auto &number_formats = stylesheet.number_formats;

        auto num_custom = std::count_if(number_formats.begin(), number_formats.end(),
            [](const number_format &nf) { return nf.get_id() >= 164; });

		serializer().start_element(xmlns, "numFmts");
		serializer().attribute("count", num_custom);

		for (const auto &num_fmt : number_formats)
		{
            if (num_fmt.get_id() < 164) continue;
			serializer().start_element(xmlns, "numFmt");
			serializer().attribute("numFmtId", num_fmt.get_id());
			serializer().attribute("formatCode", num_fmt.get_format_string());
			serializer().end_element(xmlns, "numFmt");
		}
        
		serializer().end_element(xmlns, "numFmts");
	}

	// Fonts

	if (!stylesheet.fonts.empty())
	{
		const auto &fonts = stylesheet.fonts;

		serializer().start_element(xmlns, "fonts");
		serializer().attribute("count", fonts.size());

		auto num_known_fonts = std::count_if(fonts.begin(), fonts.end(), [](const xlnt::font &f)
		{
		    static const auto known_fonts = new std::vector<xlnt::font>
		    {
			xlnt::font().color(xlnt::theme_color(1)).scheme("minor").family(2)
		    };

		    return std::find(known_fonts->begin(), known_fonts->end(), f) != known_fonts->end();
		});

		if (source_.x15_enabled())
		{
			serializer().attribute(xmlns_x14ac, "knownFonts", num_known_fonts);
		}

		for (const auto &current_font : fonts)
		{
			serializer().start_element(xmlns, "font");

			if (current_font.bold())
			{
                serializer().start_element(xmlns, "b");
                serializer().attribute("val", write_bool(true));
                serializer().end_element(xmlns, "b");
			}

			if (current_font.italic())
			{
                serializer().start_element(xmlns, "i");
                serializer().attribute("val", write_bool(true));
                serializer().end_element(xmlns, "i");
			}

			if (current_font.underlined())
			{
                serializer().start_element(xmlns, "u");
                serializer().attribute("val", current_font.underline());
                serializer().end_element(xmlns, "u");
			}

			if (current_font.strikethrough())
			{
                serializer().start_element(xmlns, "strike");
                serializer().attribute("val", write_bool(true));
                serializer().end_element(xmlns, "strike");
			}

            serializer().start_element(xmlns, "sz");
            serializer().attribute("val", current_font.size());
            serializer().end_element(xmlns, "sz");

			if (current_font.color())
			{
                serializer().start_element(xmlns, "color");
                write_color(*current_font.color());
                serializer().end_element(xmlns, "color");
			}

			serializer().start_element(xmlns, "name");
            serializer().attribute("val", current_font.name());
            serializer().end_element(xmlns, "name");

			if (current_font.family())
			{
                serializer().start_element(xmlns, "family");
                serializer().attribute("val", *current_font.family());
                serializer().end_element(xmlns, "family");
			}

			if (current_font.scheme())
			{
                serializer().start_element(xmlns, "scheme");
                serializer().attribute("val", *current_font.scheme());
                serializer().end_element(xmlns, "scheme");
			}
            
            serializer().end_element(xmlns, "font");
		}
        
		serializer().end_element(xmlns, "fonts");
	}

	// Fills

	if (!stylesheet.fills.empty())
	{
		const auto &fills = stylesheet.fills;

		serializer().start_element(xmlns, "fills");
		serializer().attribute("count", fills.size());

		for (auto &fill_ : fills)
		{
			serializer().start_element(xmlns, "fill");

			if (fill_.type() == xlnt::fill_type::pattern)
			{
				const auto &pattern = fill_.pattern_fill();

				serializer().start_element(xmlns, "patternFill");
				serializer().attribute("patternType", pattern.type());

				if (pattern.foreground())
				{
                    serializer().start_element(xmlns, "fgColor");
					write_color(*pattern.foreground());
                    serializer().end_element(xmlns, "fgColor");
				}

				if (pattern.background())
				{
                    serializer().start_element(xmlns, "bgColor");
					write_color(*pattern.background());
                    serializer().end_element(xmlns, "bgColor");
				}
                
				serializer().end_element(xmlns, "patternFill");
			}
			else if (fill_.type() == xlnt::fill_type::gradient)
			{
				const auto &gradient = fill_.gradient_fill();

				serializer().start_element(xmlns, "gradientFill");
				serializer().attribute("gradientType", gradient.type());

				if (gradient.degree() != 0.)
				{
					serializer().attribute("degree", gradient.degree());
				}

				if (gradient.left() != 0.)
				{
					serializer().attribute("left", gradient.left());
				}

				if (gradient.right() != 0.)
				{
					serializer().attribute("right", gradient.right());
				}

				if (gradient.top() != 0.)
				{
					serializer().attribute("top", gradient.top());
				}

				if (gradient.bottom() != 0.)
				{
					serializer().attribute("bottom", gradient.bottom());
				}

				for (const auto &stop : gradient.stops())
				{
					serializer().start_element(xmlns, "stop");
					serializer().attribute("position", stop.first);
					serializer().start_element(xmlns, "color");
					write_color(stop.second);
					serializer().end_element(xmlns, "color");
					serializer().end_element(xmlns, "stop");
				}
                
				serializer().end_element(xmlns, "gradientFill");
			}
            
			serializer().end_element(xmlns, "fill");
		}
        
		serializer().end_element(xmlns, "fills");
	}

	// Borders

	if (!stylesheet.borders.empty())
	{
		const auto &borders = stylesheet.borders;

		serializer().start_element(xmlns, "borders");
		serializer().attribute("count", borders.size());

		for (const auto &current_border : borders)
		{
			serializer().start_element(xmlns, "border");

			if (current_border.diagonal())
			{
				auto up = *current_border.diagonal() == diagonal_direction::both
					|| *current_border.diagonal() == diagonal_direction::up;
				serializer().attribute("diagonalUp", up ? "true" : "false");

				auto down = *current_border.diagonal() == diagonal_direction::both
					|| *current_border.diagonal() == diagonal_direction::down;
                serializer().attribute("diagonalDown", down ? "true" : "false");
			}

			for (const auto &side : xlnt::border::all_sides())
			{
				if (current_border.side(side))
				{
					const auto current_side = *current_border.side(side);

					auto side_name = to_string(side);
					serializer().start_element(xmlns, side_name);

					if (current_side.style())
					{
                        serializer().attribute("style", *current_side.style());
					}

					if (current_side.color())
					{
						serializer().start_element(xmlns, "color");
						write_color(*current_side.color());
						serializer().end_element(xmlns, "color");
					}
                    
					serializer().end_element(xmlns, side_name);
				}
			}
            
			serializer().end_element(xmlns, "border");
		}
        
		serializer().end_element(xmlns, "borders");
	}

    // Style XFs
    
    serializer().start_element(xmlns, "cellStyleXfs");
    serializer().attribute("count", stylesheet.style_impls.size());

	for (const auto &current_style_name : stylesheet.style_names)
	{
        const auto &current_style_impl = stylesheet.style_impls.at(current_style_name);

		serializer().start_element(xmlns, "xf");
        serializer().attribute("numFmtId", current_style_impl.number_format_id.get());
        serializer().attribute("fontId", current_style_impl.font_id.get());
		serializer().attribute("fillId", current_style_impl.fill_id.get());
		serializer().attribute("borderId", current_style_impl.border_id.get());

		if (current_style_impl.number_format_applied)
        {
            serializer().attribute("applyNumberFormat", write_bool(true));
        }

		if (current_style_impl.fill_applied)
        {
            serializer().attribute("applyFill", write_bool(true));
        }

		if (current_style_impl.font_applied)
        {
            serializer().attribute("applyFont", write_bool(true));
        }

		if (current_style_impl.border_applied)
        {
            serializer().attribute("applyBorder", write_bool(true));
        }

        if (current_style_impl.alignment_applied)
        {
            serializer().attribute("applyAlignment", write_bool(true));
        }

        if (current_style_impl.protection_applied)
        {
            serializer().attribute("applyProtection", write_bool(true));
        }

		if (current_style_impl.alignment_applied)
		{
            const auto &current_alignment = stylesheet.alignments[current_style_impl.alignment_id.get()];

			serializer().start_element(xmlns, "alignment");

			if (current_alignment.vertical())
			{
				serializer().attribute("vertical", current_alignment.vertical().get());
			}

			if (current_alignment.horizontal())
			{
				serializer().attribute("horizontal", current_alignment.horizontal().get());
			}

			if (current_alignment.rotation())
			{
				serializer().attribute("textRotation", current_alignment.rotation().get());
			}

			if (current_alignment.wrap())
			{
				serializer().attribute("wrapText", write_bool(current_alignment.wrap().get()));
			}

			if (current_alignment.indent())
			{
				serializer().attribute("indent", current_alignment.indent().get());
			}

			if (current_alignment.shrink())
			{
				serializer().attribute("shrinkToFit", write_bool(current_alignment.shrink().get()));
			}
            
			serializer().end_element(xmlns, "alignment");
		}

		if (current_style_impl.protection_applied)
		{
            const auto &current_protection = stylesheet.protections[current_style_impl.protection_id.get()];

			serializer().start_element(xmlns, "protection");
			serializer().attribute("locked", write_bool(current_protection.locked()));
			serializer().attribute("hidden", write_bool(current_protection.hidden()));
			serializer().end_element(xmlns, "protection");
		}
        
        serializer().end_element(xmlns, "xf");
	}
    
    serializer().end_element(xmlns, "cellStyleXfs");

	// Format XFs

	serializer().start_element(xmlns, "cellXfs");
	serializer().attribute("count", stylesheet.format_impls.size());

	for (auto &current_format_impl : stylesheet.format_impls)
	{
        serializer().start_element(xmlns, "xf");
        
        serializer().attribute("numFmtId", current_format_impl.number_format_id.get());
        serializer().attribute("fontId", current_format_impl.font_id.get());
        
        if (current_format_impl.style)
        {
            serializer().attribute("fillId", stylesheet.style_impls.at(current_format_impl.style.get()).fill_id.get());
        }
        else
        {
            serializer().attribute("fillId", current_format_impl.fill_id.get());
        }
        
		serializer().attribute("borderId", current_format_impl.border_id.get());

		if (current_format_impl.number_format_applied)
        {
            serializer().attribute("applyNumberFormat", write_bool(true));
        }

		if (current_format_impl.fill_applied)
        {
            serializer().attribute("applyFill", write_bool(true));
        }

		if (current_format_impl.font_applied)
        {
            serializer().attribute("applyFont", write_bool(true));
        }

		if (current_format_impl.border_applied)
        {
            serializer().attribute("applyBorder", write_bool(true));
        }

        if (current_format_impl.alignment_applied)
        {
            serializer().attribute("applyAlignment", write_bool(true));
        }

        if (current_format_impl.protection_applied)
        {
            serializer().attribute("applyProtection", write_bool(true));
        }

		if (current_format_impl.style.is_set())
		{
			serializer().attribute("xfId", stylesheet.style_index(current_format_impl.style.get()));
		}

		if (current_format_impl.alignment_applied)
		{
            const auto &current_alignment = stylesheet.alignments[current_format_impl.alignment_id.get()];

			serializer().start_element(xmlns, "alignment");

			if (current_alignment.vertical())
			{
				serializer().attribute("vertical", *current_alignment.vertical());
			}

			if (current_alignment.horizontal())
			{
				serializer().attribute("horizontal", *current_alignment.horizontal());
			}

			if (current_alignment.rotation())
			{
				serializer().attribute("textRotation", *current_alignment.rotation());
			}

			if (current_alignment.wrap())
			{
				serializer().attribute("wrapText", write_bool(*current_alignment.wrap()));
			}

			if (current_alignment.indent())
			{
				serializer().attribute("indent", *current_alignment.indent());
			}

			if (current_alignment.shrink())
			{
				serializer().attribute("shrinkToFit", write_bool(*current_alignment.shrink()));
			}
            
			serializer().end_element(xmlns, "alignment");
		}

		if (current_format_impl.protection_applied)
		{
            const auto &current_protection = stylesheet.protections[current_format_impl.protection_id.get()];

			serializer().start_element(xmlns, "protection");
			serializer().attribute("locked", write_bool(current_protection.locked()));
			serializer().attribute("hidden", write_bool(current_protection.hidden()));
			serializer().end_element(xmlns, "protection");
		}

        serializer().end_element(xmlns, "xf");
	}
    
    serializer().end_element(xmlns, "cellXfs");

	// Styles

	serializer().start_element(xmlns, "cellStyles");
    serializer().attribute("count", stylesheet.style_impls.size());
	std::size_t style_index = 0;

	for (auto &current_style_name : stylesheet.style_names)
	{
        const auto &current_style = stylesheet.style_impls.at(current_style_name);

		serializer().start_element(xmlns, "cellStyle");

		serializer().attribute("name", current_style.name);
		serializer().attribute("xfId", style_index++);

		if (current_style.builtin_id)
		{
			serializer().attribute("builtinId", current_style.builtin_id.get());
		}

		if (current_style.hidden_style)
		{
			serializer().attribute("hidden", write_bool(true));
		}

		if (current_style.custom_builtin)
		{
			serializer().attribute("customBuiltin", write_bool(current_style.custom_builtin.get()));
		}
        
		serializer().end_element(xmlns, "cellStyle");
	}
    
    serializer().end_element(xmlns, "cellStyles");

    serializer().start_element(xmlns, "dxfs");
	serializer().attribute("count", "0");
    serializer().end_element(xmlns, "dxfs");
    
    serializer().start_element(xmlns, "tableStyles");
	serializer().attribute("count", "0");
	serializer().attribute("defaultTableStyle", "TableStyleMedium9");
	serializer().attribute("defaultPivotStyle", "PivotStyleMedium7");
    serializer().end_element(xmlns, "tableStyles");
    
    if (!stylesheet.colors.empty())
    {
        serializer().start_element(xmlns, "indexedColors");

        for (auto &c : stylesheet.colors)
        {
            serializer().start_element(xmlns, "rgbColor");
            serializer().attribute("rgb", c.get_rgb().get_hex_string());
            serializer().end_element(xmlns, "rgbColor");
        }

        serializer().end_element(xmlns, "indexedColors");
    }

	serializer().start_element(xmlns, "extLst");
	serializer().start_element(xmlns, "ext");
    serializer().namespace_decl(xmlns_x14, "x14");
	serializer().attribute("uri", "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}");
	serializer().start_element(xmlns_x14, "slicerStyles");
	serializer().attribute("defaultSlicerStyle", "SlicerStyleLight1");
	serializer().end_element(xmlns_x14, "slicerStyles");
	serializer().end_element(xmlns, "ext");
    serializer().end_element(xmlns, "extLst");
    
    serializer().end_element(xmlns, "styleSheet");
}

void xlsx_producer::write_theme(const relationship &/*rel*/)
{
    static const auto &xmlns_a = constants::get_namespace("drawingml");
    static const auto &xmlns_thm15 = constants::get_namespace("thm15");
    
	serializer().start_element(xmlns_a, "theme");
    serializer().namespace_decl(xmlns_a, "a");
    serializer().attribute("name", "Office Theme");

	serializer().start_element(xmlns_a, "themeElements");
	serializer().start_element(xmlns_a, "clrScheme");
	serializer().attribute("name", "Office");

	struct scheme_element
	{
		std::string name;
		std::string sub_element_name;
		std::string val;
	};

	std::vector<scheme_element> scheme_elements = {
		{ "dk1", "sysClr", "windowText" },
        { "lt1", "sysClr", "window" },
		{ "dk2", "srgbClr", "44546A" },
        { "lt2", "srgbClr", "E7E6E6" },
		{ "accent1", "srgbClr", "5B9BD5" },
        { "accent2", "srgbClr", "ED7D31" },
		{ "accent3", "srgbClr", "A5A5A5" },
        { "accent4", "srgbClr", "FFC000" },
		{ "accent5", "srgbClr", "4472C4" },
        { "accent6", "srgbClr", "70AD47" },
		{ "hlink", "srgbClr", "0563C1" },
        { "folHlink", "srgbClr", "954F72" },
	};

	for (auto element : scheme_elements)
	{
        serializer().start_element(xmlns_a, element.name);
        serializer().start_element(xmlns_a, element.sub_element_name);
        serializer().attribute("val", element.val);

		if (element.name == "dk1")
		{
            serializer().attribute("lastClr", "000000");
		}
		else if (element.name == "lt1")
		{
            serializer().attribute("lastClr", "FFFFFF");
		}
        
        serializer().end_element(xmlns_a, element.sub_element_name);
        serializer().end_element(xmlns_a, element.name);
	}
    
    serializer().end_element(xmlns_a, "clrScheme");

	struct font_scheme
	{
		bool typeface;
		std::string script;
		std::string major;
		std::string minor;
	};

	static const auto font_schemes = new std::vector<font_scheme> {
		{ true, "latin", "Calibri Light", "Calibri" },
		{ true, "ea", "", "" },
		{ true, "cs", "", "" },
		{ false, "Jpan", "Yu Gothic Light", "Yu Gothic" },
		{ false, "Hang", "\xeb\xa7\x91\xec\x9d\x80 \xea\xb3\xa0\xeb\x94\x95",
		"\xeb\xa7\x91\xec\x9d\x80 \xea\xb3\xa0\xeb\x94\x95" },
		{ false, "Hans", "DengXian Light", "DengXian" },
		{ false, "Hant", "\xe6\x96\xb0\xe7\xb4\xb0\xe6\x98\x8e\xe9\xab\x94",
		"\xe6\x96\xb0\xe7\xb4\xb0\xe6\x98\x8e\xe9\xab\x94" },
		{ false, "Arab", "Times New Roman", "Arial" },
		{ false, "Hebr", "Times New Roman", "Arial" },
		{ false, "Thai", "Tahoma", "Tahoma" },
		{ false, "Ethi", "Nyala", "Nyala" },
		{ false, "Beng", "Vrinda", "Vrinda" },
		{ false, "Gujr", "Shruti", "Shruti" },
		{ false, "Khmr", "MoolBoran", "DaunPenh" },
		{ false, "Knda", "Tunga", "Tunga" },
		{ false, "Guru", "Raavi", "Raavi" },
		{ false, "Cans", "Euphemia", "Euphemia" },
		{ false, "Cher", "Plantagenet Cherokee", "Plantagenet Cherokee" },
		{ false, "Yiii", "Microsoft Yi Baiti", "Microsoft Yi Baiti" },
		{ false, "Tibt", "Microsoft Himalaya", "Microsoft Himalaya" },
		{ false, "Thaa", "MV Boli", "MV Boli" },
		{ false, "Deva", "Mangal", "Mangal" },
		{ false, "Telu", "Gautami", "Gautami" },
		{ false, "Taml", "Latha", "Latha" },
		{ false, "Syrc", "Estrangelo Edessa", "Estrangelo Edessa" },
		{ false, "Orya", "Kalinga", "Kalinga" },
		{ false, "Mlym", "Kartika", "Kartika" },
		{ false, "Laoo", "DokChampa", "DokChampa" },
		{ false, "Sinh", "Iskoola Pota", "Iskoola Pota" },
		{ false, "Mong", "Mongolian Baiti", "Mongolian Baiti" },
		{ false, "Viet", "Times New Roman", "Arial" },
		{ false, "Uigh", "Microsoft Uighur", "Microsoft Uighur" },
		{ false, "Geor", "Sylfaen", "Sylfaen" }
	};

	serializer().start_element(xmlns_a, "fontScheme");
    serializer().attribute("name", "Office");

    for (auto major : {true, false})
    {
        serializer().start_element(xmlns_a, major ? "majorFont" : "minorFont");
        
        for (const auto &scheme : *font_schemes)
        {
            const auto scheme_value = major ? scheme.major : scheme.minor;

            if (scheme.typeface)
            {
                serializer().start_element(xmlns_a, scheme.script);
                serializer().attribute("typeface", scheme_value);

                if (scheme_value == "Calibri Light")
                {
                    serializer().attribute("panose", "020F0302020204030204");
                }
                else if (scheme_value == "Calibri")
                {
                    serializer().attribute("panose", "020F0502020204030204");
                }
                
                serializer().end_element(xmlns_a, scheme.script);
            }
            else
            {
                serializer().start_element(xmlns_a, "font");
                serializer().attribute("script", scheme.script);
                serializer().attribute("typeface", scheme_value);
                serializer().end_element(xmlns_a, "font");
            }
        }
        
        serializer().end_element(xmlns_a, major ? "majorFont" : "minorFont");
    }
    
	serializer().end_element(xmlns_a, "fontScheme");

	serializer().start_element(xmlns_a, "fmtScheme");
    serializer().attribute("name", "Office");

	serializer().start_element(xmlns_a, "fillStyleLst");
	serializer().start_element(xmlns_a, "solidFill");
    serializer().start_element(xmlns_a, "schemeClr");
    serializer().attribute("val", "phClr");
    serializer().end_element(xmlns_a, "schemeClr");
	serializer().end_element(xmlns_a, "solidFill");

	serializer().start_element(xmlns_a, "gradFill");
	serializer().attribute("rotWithShape", "1");
	serializer().start_element(xmlns_a, "gsLst");
    
	serializer().start_element(xmlns_a, "gs");
	serializer().attribute("pos", "0");
	serializer().start_element(xmlns_a, "schemeClr");
	serializer().attribute("val", "phClr");
	serializer().start_element(xmlns_a, "lumMod");
    serializer().attribute("val", "110000");
	serializer().end_element(xmlns_a, "lumMod");
	serializer().start_element(xmlns_a, "satMod");
    serializer().attribute("val", "105000");
    serializer().end_element(xmlns_a, "satMod");
	serializer().start_element(xmlns_a, "tint");
    serializer().attribute("val", "67000");
	serializer().end_element(xmlns_a, "tint");
	serializer().end_element(xmlns_a, "schemeClr");
	serializer().end_element(xmlns_a, "gs");

	serializer().start_element(xmlns_a, "gs");
	serializer().attribute("pos", "50000");
	serializer().start_element(xmlns_a, "schemeClr");
	serializer().attribute("val", "phClr");
	serializer().start_element(xmlns_a, "lumMod");
    serializer().attribute("val", "105000");
	serializer().end_element(xmlns_a, "lumMod");
	serializer().start_element(xmlns_a, "satMod");
    serializer().attribute("val", "103000");
    serializer().end_element(xmlns_a, "satMod");
	serializer().start_element(xmlns_a, "tint");
    serializer().attribute("val", "73000");
	serializer().end_element(xmlns_a, "tint");
	serializer().end_element(xmlns_a, "schemeClr");
	serializer().end_element(xmlns_a, "gs");

	serializer().start_element(xmlns_a, "gs");
	serializer().attribute("pos", "100000");
	serializer().start_element(xmlns_a, "schemeClr");
	serializer().attribute("val", "phClr");
	serializer().start_element(xmlns_a, "lumMod");
    serializer().attribute("val", "105000");
	serializer().end_element(xmlns_a, "lumMod");
	serializer().start_element(xmlns_a, "satMod");
    serializer().attribute("val", "109000");
    serializer().end_element(xmlns_a, "satMod");
	serializer().start_element(xmlns_a, "tint");
    serializer().attribute("val", "81000");
	serializer().end_element(xmlns_a, "tint");
	serializer().end_element(xmlns_a, "schemeClr");
	serializer().end_element(xmlns_a, "gs");
    
	serializer().end_element(xmlns_a, "gsLst");

	serializer().start_element(xmlns_a, "lin");
	serializer().attribute("ang", "5400000");
	serializer().attribute("scaled", "0");
	serializer().end_element(xmlns_a, "lin");
    
	serializer().end_element(xmlns_a, "gradFill");

	serializer().start_element(xmlns_a, "gradFill");
	serializer().attribute("rotWithShape", "1");
	serializer().start_element(xmlns_a, "gsLst");

	serializer().start_element(xmlns_a, "gs");
	serializer().attribute("pos", "0");
	serializer().start_element(xmlns_a, "schemeClr");
	serializer().attribute("val", "phClr");
	serializer().start_element(xmlns_a, "satMod");
    serializer().attribute("val", "103000");
	serializer().end_element(xmlns_a, "satMod");
	serializer().start_element(xmlns_a, "lumMod");
    serializer().attribute("val", "102000");
	serializer().end_element(xmlns_a, "lumMod");
	serializer().start_element(xmlns_a, "tint");
    serializer().attribute("val", "94000");
	serializer().end_element(xmlns_a, "tint");
	serializer().end_element(xmlns_a, "schemeClr");
    serializer().end_element(xmlns_a, "gs");

	serializer().start_element(xmlns_a, "gs");
	serializer().attribute("pos", "50000");
	serializer().start_element(xmlns_a, "schemeClr");
	serializer().attribute("val", "phClr");
	serializer().start_element(xmlns_a, "satMod");
    serializer().attribute("val", "110000");
	serializer().end_element(xmlns_a, "satMod");
	serializer().start_element(xmlns_a, "lumMod");
    serializer().attribute("val", "100000");
	serializer().end_element(xmlns_a, "lumMod");
	serializer().start_element(xmlns_a, "shade");
    serializer().attribute("val", "100000");
	serializer().end_element(xmlns_a, "shade");
	serializer().end_element(xmlns_a, "schemeClr");
    serializer().end_element(xmlns_a, "gs");

	serializer().start_element(xmlns_a, "gs");
	serializer().attribute("pos", "100000");
	serializer().start_element(xmlns_a, "schemeClr");
	serializer().attribute("val", "phClr");
	serializer().start_element(xmlns_a, "lumMod");
    serializer().attribute("val", "99000");
	serializer().end_element(xmlns_a, "lumMod");
	serializer().start_element(xmlns_a, "satMod");
    serializer().attribute("val", "120000");
	serializer().end_element(xmlns_a, "satMod");
	serializer().start_element(xmlns_a, "shade");
    serializer().attribute("val", "78000");
	serializer().end_element(xmlns_a, "shade");
	serializer().end_element(xmlns_a, "schemeClr");
    serializer().end_element(xmlns_a, "gs");
    
	serializer().end_element(xmlns_a, "gsLst");

	serializer().start_element(xmlns_a, "lin");
	serializer().attribute("ang", "5400000");
	serializer().attribute("scaled", "0");
    serializer().end_element(xmlns_a, "lin");
    
    serializer().end_element(xmlns_a, "gradFill");
	serializer().end_element(xmlns_a, "fillStyleLst");

	serializer().start_element(xmlns_a, "lnStyleLst");

	serializer().start_element(xmlns_a, "ln");
	serializer().attribute("w", "6350");
	serializer().attribute("cap", "flat");
	serializer().attribute("cmpd", "sng");
	serializer().attribute("algn", "ctr");
	serializer().start_element(xmlns_a, "solidFill");
	serializer().start_element(xmlns_a, "schemeClr");
	serializer().attribute("val", "phClr");
	serializer().end_element(xmlns_a, "schemeClr");
	serializer().end_element(xmlns_a, "solidFill");
	serializer().start_element(xmlns_a, "prstDash");
    serializer().attribute("val", "solid");
	serializer().end_element(xmlns_a, "prstDash");
	serializer().start_element(xmlns_a, "miter");
    serializer().attribute("lim", "800000");
	serializer().end_element(xmlns_a, "miter");
	serializer().end_element(xmlns_a, "ln");

	serializer().start_element(xmlns_a, "ln");
	serializer().attribute("w", "12700");
	serializer().attribute("cap", "flat");
	serializer().attribute("cmpd", "sng");
	serializer().attribute("algn", "ctr");
	serializer().start_element(xmlns_a, "solidFill");
	serializer().start_element(xmlns_a, "schemeClr");
	serializer().attribute("val", "phClr");
	serializer().end_element(xmlns_a, "schemeClr");
	serializer().end_element(xmlns_a, "solidFill");
	serializer().start_element(xmlns_a, "prstDash");
    serializer().attribute("val", "solid");
	serializer().end_element(xmlns_a, "prstDash");
	serializer().start_element(xmlns_a, "miter");
    serializer().attribute("lim", "800000");
	serializer().end_element(xmlns_a, "miter");
	serializer().end_element(xmlns_a, "ln");

	serializer().start_element(xmlns_a, "ln");
	serializer().attribute("w", "19050");
	serializer().attribute("cap", "flat");
	serializer().attribute("cmpd", "sng");
	serializer().attribute("algn", "ctr");
	serializer().start_element(xmlns_a, "solidFill");
	serializer().start_element(xmlns_a, "schemeClr");
	serializer().attribute("val", "phClr");
	serializer().end_element(xmlns_a, "schemeClr");
	serializer().end_element(xmlns_a, "solidFill");
	serializer().start_element(xmlns_a, "prstDash");
    serializer().attribute("val", "solid");
	serializer().end_element(xmlns_a, "prstDash");
	serializer().start_element(xmlns_a, "miter");
    serializer().attribute("lim", "800000");
	serializer().end_element(xmlns_a, "miter");
	serializer().end_element(xmlns_a, "ln");
    
	serializer().end_element(xmlns_a, "lnStyleLst");

	serializer().start_element(xmlns_a, "effectStyleLst");

	serializer().start_element(xmlns_a, "effectStyle");
	serializer().element(xmlns_a, "effectLst", "");
	serializer().end_element(xmlns_a, "effectStyle");

	serializer().start_element(xmlns_a, "effectStyle");
	serializer().element(xmlns_a, "effectLst", "");
	serializer().end_element(xmlns_a, "effectStyle");

	serializer().start_element(xmlns_a, "effectStyle");
	serializer().start_element(xmlns_a, "effectLst");
	serializer().start_element(xmlns_a, "outerShdw");
	serializer().attribute("blurRad", "57150");
	serializer().attribute("dist", "19050");
	serializer().attribute("dir", "5400000");
	serializer().attribute("algn", "ctr");
    serializer().attribute("rotWithShape", "0");
	serializer().start_element(xmlns_a, "srgbClr");
	serializer().attribute("val", "000000");
	serializer().start_element(xmlns_a, "alpha");
    serializer().attribute("val", "63000");
	serializer().end_element(xmlns_a, "alpha");
	serializer().end_element(xmlns_a, "srgbClr");
	serializer().end_element(xmlns_a, "outerShdw");
	serializer().end_element(xmlns_a, "effectLst");
	serializer().end_element(xmlns_a, "effectStyle");
    
	serializer().end_element(xmlns_a, "effectStyleLst");

	serializer().start_element(xmlns_a, "bgFillStyleLst");

	serializer().start_element(xmlns_a, "solidFill");
    serializer().start_element(xmlns_a, "schemeClr");
    serializer().attribute("val", "phClr");
	serializer().end_element(xmlns_a, "schemeClr");
	serializer().end_element(xmlns_a, "solidFill");

	serializer().start_element(xmlns_a, "solidFill");
    serializer().start_element(xmlns_a, "schemeClr");
    serializer().attribute("val", "phClr");
	serializer().start_element(xmlns_a, "tint");
    serializer().attribute("val", "95000");
	serializer().end_element(xmlns_a, "tint");
	serializer().start_element(xmlns_a, "satMod");
    serializer().attribute("val", "170000");
	serializer().end_element(xmlns_a, "satMod");
	serializer().end_element(xmlns_a, "schemeClr");
	serializer().end_element(xmlns_a, "solidFill");

	serializer().start_element(xmlns_a, "gradFill");
	serializer().attribute("rotWithShape", "1");
	serializer().start_element(xmlns_a, "gsLst");

	serializer().start_element(xmlns_a, "gs");
	serializer().attribute("pos", "0");
	serializer().start_element(xmlns_a, "schemeClr");
	serializer().attribute("val", "phClr");
	serializer().start_element(xmlns_a, "tint");
    serializer().attribute("val", "93000");
	serializer().end_element(xmlns_a, "tint");
	serializer().start_element(xmlns_a, "satMod");
    serializer().attribute("val", "150000");
	serializer().end_element(xmlns_a, "satMod");
	serializer().start_element(xmlns_a, "shade");
    serializer().attribute("val", "98000");
	serializer().end_element(xmlns_a, "shade");
	serializer().start_element(xmlns_a, "lumMod");
    serializer().attribute("val", "102000");
	serializer().end_element(xmlns_a, "lumMod");
	serializer().end_element(xmlns_a, "schemeClr");
	serializer().end_element(xmlns_a, "gs");

	serializer().start_element(xmlns_a, "gs");
	serializer().attribute("pos", "50000");
	serializer().start_element(xmlns_a, "schemeClr");
	serializer().attribute("val", "phClr");
	serializer().start_element(xmlns_a, "tint");
    serializer().attribute("val", "98000");
	serializer().end_element(xmlns_a, "tint");
	serializer().start_element(xmlns_a, "satMod");
    serializer().attribute("val", "130000");
	serializer().end_element(xmlns_a, "satMod");
	serializer().start_element(xmlns_a, "shade");
    serializer().attribute("val", "90000");
	serializer().end_element(xmlns_a, "shade");
	serializer().start_element(xmlns_a, "lumMod");
    serializer().attribute("val", "103000");
	serializer().end_element(xmlns_a, "lumMod");
	serializer().end_element(xmlns_a, "schemeClr");
	serializer().end_element(xmlns_a, "gs");

	serializer().start_element(xmlns_a, "gs");
	serializer().attribute("pos", "100000");
	serializer().start_element(xmlns_a, "schemeClr");
	serializer().attribute("val", "phClr");
	serializer().start_element(xmlns_a, "shade");
    serializer().attribute("val", "63000");
	serializer().end_element(xmlns_a, "shade");
	serializer().start_element(xmlns_a, "satMod");
    serializer().attribute("val", "120000");
	serializer().end_element(xmlns_a, "satMod");
	serializer().end_element(xmlns_a, "schemeClr");
	serializer().end_element(xmlns_a, "gs");

	serializer().end_element(xmlns_a, "gsLst");

	serializer().start_element(xmlns_a, "lin");
	serializer().attribute("ang", "5400000");
	serializer().attribute("scaled", "0");
	serializer().end_element(xmlns_a, "lin");

	serializer().end_element(xmlns_a, "gradFill");

	serializer().end_element(xmlns_a, "bgFillStyleLst");
	serializer().end_element(xmlns_a, "fmtScheme");
	serializer().end_element(xmlns_a, "themeElements");

	serializer().element(xmlns_a, "objectDefaults", "");
	serializer().element(xmlns_a, "extraClrSchemeLst", "");

	serializer().start_element(xmlns_a, "extLst");
	serializer().start_element(xmlns_a, "ext");
	serializer().attribute("uri", "{05A4C25C-085E-4340-85A3-A5531E510DB2}");
    serializer().start_element(xmlns_thm15, "themeFamily");
    serializer().namespace_decl(xmlns_thm15, "thm15");
	serializer().attribute("name", "Office Theme");
	serializer().attribute("id", "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}");
	serializer().attribute("vid", "{4A3C46E8-61CC-4603-A589-7422A47A8E4A}");
	serializer().end_element(xmlns_thm15, "themeFamily");
	serializer().end_element(xmlns_a, "ext");
	serializer().end_element(xmlns_a, "extLst");
    
    serializer().end_element(xmlns_a, "theme");
}

void xlsx_producer::write_volatile_dependencies(const relationship &/*rel*/)
{
	serializer().start_element("volTypes");
}

void xlsx_producer::write_worksheet(const relationship &rel)
{
	static const auto &xmlns = constants::get_namespace("worksheet");
	static const auto &xmlns_r = constants::get_namespace("r");
	static const auto &xmlns_mc = constants::get_namespace("mc");
	static const auto &xmlns_x14ac = constants::get_namespace("x14ac");

	auto worksheet_part = rel.get_source().get_path().parent().append(rel.get_target().get_path());
	auto worksheet_rels = source_.get_manifest().get_relationships(worksheet_part);

	auto title = std::find_if(source_.d_->sheet_title_rel_id_map_.begin(),
		source_.d_->sheet_title_rel_id_map_.end(),
		[&](const std::pair<std::string, std::string> &p)
	{
		return p.second == rel.get_id();
	})->first;

	auto ws = source_.get_sheet_by_title(title);

	serializer().start_element(xmlns, "worksheet");
    serializer().namespace_decl(xmlns, "");
    serializer().namespace_decl(xmlns_r, "r");

	if (ws.x14ac_enabled())
	{
		serializer().namespace_decl(xmlns_mc, "mc");
		serializer().namespace_decl(xmlns_x14ac, "x14ac");
		serializer().attribute(xmlns_mc, "Ignorable", "x14ac");
	}

	if (ws.has_page_setup())
	{
		serializer().start_element(xmlns, "sheetPr");

		serializer().start_element(xmlns, "outlinePr");
        serializer().attribute("summaryBelow", "1");
		serializer().attribute("summaryRight", "1");
        serializer().end_element(xmlns, "outlinePr");

		serializer().start_element(xmlns, "pageSetUpPr");
		serializer().attribute("fitToPage", write_bool(ws.get_page_setup().fit_to_page()));
		serializer().end_element(xmlns, "pageSetUpPr");

		serializer().end_element(xmlns, "sheetPr");
	}

	if (ws.has_dimension())
	{
		serializer().start_element(xmlns, "dimension");
        const auto dimension = ws.calculate_dimension();
		serializer().attribute("ref", dimension.is_single_cell()
            ? dimension.get_top_left().to_string() : dimension.to_string());
		serializer().end_element(xmlns, "dimension");
	}

	if (ws.has_view())
	{
		serializer().start_element(xmlns, "sheetViews");
		serializer().start_element(xmlns, "sheetView");

		const auto view = ws.get_view();

		serializer().attribute("tabSelected", "1");
		serializer().attribute("workbookViewId", "0");

		if (!view.get_selections().empty() && !ws.has_frozen_panes())
		{
			serializer().start_element(xmlns, "selection");

			const auto &first_selection = view.get_selections().front();

			if (first_selection.has_active_cell())
			{
				auto active_cell = first_selection.get_active_cell();
				serializer().attribute("activeCell", active_cell.to_string());
				serializer().attribute("sqref", active_cell.to_string());
			}
            
            serializer().end_element(xmlns, "selection");
		}

		auto active_pane = "bottomRight"s;

		if (ws.has_frozen_panes())
		{
			serializer().start_element(xmlns, "pane");

			if (ws.get_frozen_panes().get_column_index() > 1)
			{
				serializer().attribute("xSplit", ws.get_frozen_panes().get_column_index().index - 1);
				active_pane = "topRight";
			}

			if (ws.get_frozen_panes().get_row() > 1)
			{
				serializer().attribute("ySplit", ws.get_frozen_panes().get_row() - 1);
				active_pane = "bottomLeft";
			}

/*
			if (ws.get_frozen_panes().get_row() > 1 && ws.get_frozen_panes().get_column_index() > 1)
			{
				serializer().start_element(xmlns, "selection");
				serializer().attribute("pane", "topRight");
				auto bottom_left_node = sheet_view_node.append_child("selection");
				serializer().attribute("pane", "bottomLeft");
				active_pane = "bottomRight";
			}
            */

			serializer().attribute("topLeftCell", ws.get_frozen_panes().to_string());
			serializer().attribute("activePane", active_pane);
			serializer().attribute("state", "frozen");

			serializer().start_element(xmlns, "selection");

			if (ws.get_frozen_panes().get_row() > 1 && ws.get_frozen_panes().get_column_index() > 1)
			{
				serializer().attribute("pane", "bottomRight");
			}
			else if (ws.get_frozen_panes().get_row() > 1)
			{
				serializer().attribute("pane", "bottomLeft");
			}
			else if (ws.get_frozen_panes().get_column_index() > 1)
			{
				serializer().attribute("pane", "topRight");
			}
            
            serializer().end_element(xmlns, "selection");
            serializer().end_element(xmlns, "pane");
		}

		serializer().end_element(xmlns, "sheetView");
        serializer().end_element(xmlns, "sheetViews");
	}

	if (ws.has_format_properties())
	{
		serializer().start_element(xmlns, "sheetFormatPr");
		serializer().attribute("baseColWidth", "10");
		serializer().attribute("defaultRowHeight", "16");
        if (ws.x14ac_enabled())
        {
            serializer().attribute(xmlns_x14ac, "dyDescent", "0.2");
        }
		serializer().end_element(xmlns, "sheetFormatPr");
	}

	bool has_column_properties = false;

	for (auto column = ws.get_lowest_column(); column <= ws.get_highest_column(); column++)
	{
		if (ws.has_column_properties(column))
		{
			has_column_properties = true;
			break;
		}
	}

	if (has_column_properties)
	{
		serializer().start_element(xmlns, "cols");

		for (auto column = ws.get_lowest_column(); column <= ws.get_highest_column(); column++)
		{
			if (!ws.has_column_properties(column)) continue;

			const auto &props = ws.get_column_properties(column);

            serializer().start_element(xmlns, "col");
			serializer().attribute("min", column.index);
			serializer().attribute("max", column.index);
			serializer().attribute("width", props.width);
			serializer().attribute("style", props.style);
			serializer().attribute("customWidth", write_bool(props.custom));
            serializer().end_element(xmlns, "col");
		}
        
		serializer().end_element(xmlns, "cols");
	}

	std::unordered_map<std::string, std::string> hyperlink_references;

	serializer().start_element(xmlns, "sheetData");
	const auto &shared_strings = ws.get_workbook().get_shared_strings();
	std::vector<cell_reference> cells_with_comments;

	for (auto row : ws.rows())
	{
		auto min = static_cast<xlnt::row_t>(row.num_cells());
		xlnt::row_t max = 0;
		bool any_non_null = false;

		for (auto cell : row)
		{
			min = std::min(min, cell.get_column().index);
			max = std::max(max, cell.get_column().index);

			if (!cell.garbage_collectible())
			{
				any_non_null = true;
			}
		}

		if (!any_non_null)
		{
			continue;
		}

		serializer().start_element(xmlns, "row");

		serializer().attribute("r", row.front().get_row());
		serializer().attribute("spans", std::to_string(min) + ":" + std::to_string(max));

		if (ws.has_row_properties(row.front().get_row()))
		{
			serializer().attribute("customHeight", "1");
			auto height = ws.get_row_properties(row.front().get_row()).height;

			if (std::fabs(height - std::floor(height)) == 0.L)
			{
				serializer().attribute("ht", std::to_string(static_cast<long long int>(height)) + ".0");
			}
			else
			{
				serializer().attribute("ht", height);
			}
		}

        if (!skip_unknown_elements)
        {
            serializer().attribute(xmlns_x14ac, "dyDescent", 0.25);
        }

		for (auto cell : row)
		{
			if (cell.has_comment())
			{
				cells_with_comments.push_back(cell.get_reference());
			}

			if (!cell.garbage_collectible())
			{
				serializer().start_element(xmlns, "c");
				serializer().attribute("r", cell.get_reference().to_string());
            
				if (cell.has_format())
				{
					serializer().attribute("s", cell.format().id());
				}

				if (cell.get_data_type() == cell::type::string)
				{
					if (cell.has_formula())
					{
						serializer().attribute("t", "str");
						serializer().element(xmlns, "f", cell.get_formula());
						serializer().element(xmlns, "v", cell.to_string());
                        serializer().end_element(xmlns, "c");

						continue;
					}

					int match_index = -1;

					for (std::size_t i = 0; i < shared_strings.size(); i++)
					{
						if (shared_strings[i] == cell.get_value<formatted_text>())
						{
							match_index = static_cast<int>(i);
							break;
						}
					}

					if (match_index == -1)
					{
						if (cell.get_value<std::string>().empty())
						{
							serializer().attribute("t", "s");
						}
						else
						{
							serializer().attribute("t", "inlineStr");
							serializer().start_element(xmlns, "is");
							serializer().element(xmlns, "t", cell.get_value<std::string>());
							serializer().end_element(xmlns, "is");
						}
					}
					else
					{
						serializer().attribute("t", "s");
						serializer().element(xmlns, "v", match_index);
					}
				}
				else
				{
					if (cell.get_data_type() != cell::type::null)
					{
						if (cell.get_data_type() == cell::type::boolean)
						{
							serializer().attribute("t", "b");
							serializer().element(xmlns, "v", write_bool(cell.get_value<bool>()));
						}
						else if (cell.get_data_type() == cell::type::numeric)
						{
							if (cell.has_formula())
							{
								serializer().element(xmlns, "f", cell.get_formula());
								serializer().element(xmlns, "v", cell.to_string());
                                serializer().end_element(xmlns, "c");

								continue;
							}

							serializer().attribute("t", "n");
							serializer().start_element(xmlns, "v");

							if (is_integral(cell.get_value<long double>()))
							{
                                serializer().characters(cell.get_value<long long>());
							}
							else
							{
								std::stringstream ss;
								ss.precision(20);
								ss << cell.get_value<long double>();
								ss.str();
                                serializer().characters(ss.str());
							}
                            
                            serializer().end_element(xmlns, "v");
						}
					}
					else if (cell.has_formula())
					{
                        serializer().element(xmlns, "f", cell.get_formula());
                        // todo (but probably not) could calculate the formula and set the value here
                        serializer().end_element(xmlns, "c");

                        continue;
					}
				}
                
                serializer().end_element(xmlns, "c");
			}
		}
        
        serializer().end_element(xmlns, "row");
	}

    serializer().end_element(xmlns, "sheetData");

	if (ws.has_auto_filter())
	{
		serializer().start_element(xmlns, "autoFilter");
		serializer().attribute("ref", ws.get_auto_filter().to_string());
		serializer().end_element(xmlns, "autoFilter");
	}

	if (!ws.get_merged_ranges().empty())
	{
		serializer().start_element(xmlns, "mergeCells");
		serializer().attribute("count", ws.get_merged_ranges().size());

		for (auto merged_range : ws.get_merged_ranges())
		{
            serializer().start_element(xmlns, "mergeCell");
			serializer().attribute("ref", merged_range.to_string());
            serializer().end_element(xmlns, "mergeCell");
		}
        
		serializer().end_element(xmlns, "mergeCells");
	}

	if (!worksheet_rels.empty())
	{
		std::vector<relationship> hyperlink_sheet_rels;

		for (const auto &sheet_rel : worksheet_rels)
		{
			if (sheet_rel.get_type() == relationship::type::hyperlink)
			{
				hyperlink_sheet_rels.push_back(sheet_rel);
			}
		}

		if (!hyperlink_sheet_rels.empty())
		{
            serializer().start_element(xmlns, "hyperlinks");

			for (const auto &hyperlink_rel : hyperlink_sheet_rels)
			{
                serializer().start_element(xmlns, "hyperlink");
				serializer().attribute("display", hyperlink_rel.get_target().get_path().string());
				serializer().attribute("ref", hyperlink_references.at(hyperlink_rel.get_id()));
				serializer().attribute(xmlns_r, "id", hyperlink_rel.get_id());
                serializer().end_element(xmlns, "hyperlink");
			}
            
            serializer().end_element(xmlns, "hyperlinks");
		}
	}

	if (ws.has_page_setup())
	{
		serializer().start_element(xmlns, "printOptions");
		serializer().attribute("horizontalCentered", write_bool(ws.get_page_setup().get_horizontal_centered()));
		serializer().attribute("verticalCentered", write_bool(ws.get_page_setup().get_vertical_centered()));
		serializer().end_element(xmlns, "printOptions");
	}

	if (ws.has_page_margins())
	{
		serializer().start_element(xmlns, "pageMargins");

		//TODO: there must be a better way to do this
		auto remove_trailing_zeros = [](const std::string &n)
		{
			auto decimal = n.find('.');

			if (decimal == std::string::npos) return n;

			auto index = n.size() - 1;

			while (index >= decimal && n[index] == '0')
			{
				index--;
			}

			if (index == decimal)
			{
				return n.substr(0, decimal);
			}

			return n.substr(0, index + 1);
		};

		serializer().attribute("left", remove_trailing_zeros(std::to_string(ws.get_page_margins().get_left())));
		serializer().attribute("right", remove_trailing_zeros(std::to_string(ws.get_page_margins().get_right())));
		serializer().attribute("top", remove_trailing_zeros(std::to_string(ws.get_page_margins().get_top())));
		serializer().attribute("bottom", remove_trailing_zeros(std::to_string(ws.get_page_margins().get_bottom())));
		serializer().attribute("header", remove_trailing_zeros(std::to_string(ws.get_page_margins().get_header())));
		serializer().attribute("footer", remove_trailing_zeros(std::to_string(ws.get_page_margins().get_footer())));
        
        serializer().end_element(xmlns, "pageMargins");
	}

	if (ws.has_page_setup())
	{
		serializer().start_element(xmlns, "pageSetup");
		serializer().attribute("orientation",
            ws.get_page_setup().get_orientation() == xlnt::orientation::landscape
            ? "landscape" : "portrait");
		serializer().attribute("paperSize",  static_cast<std::size_t>(ws.get_page_setup().get_paper_size()));
        serializer().attribute("fitToHeight", write_bool(ws.get_page_setup().fit_to_height()));
		serializer().attribute("fitToWidth", write_bool(ws.get_page_setup().fit_to_width()));
		serializer().end_element(xmlns, "pageSetup");
	}

	if (!ws.get_header_footer().is_default())
	{
        // todo: this shouldn't be hardcoded
	    static const auto header_text = new std::string(
			"&L&\"Calibri,Regular\"&K000000Left Header Text&C&\"Arial,Regular\"&6&K445566Center Header "
			"Text&R&\"Arial,Bold\"&8&K112233Right Header Text");
	    static const auto footer_text = new std::string(
			"&L&\"Times New Roman,Regular\"&10&K445566Left Footer Text_x000D_And &D and &T&C&\"Times New "
			"Roman,Bold\"&12&K778899Center Footer Text &Z&F on &A&R&\"Times New Roman,Italic\"&14&KAABBCCRight Footer "
			"Text &P of &N");

		serializer().start_element(xmlns, "headerFooter");
		serializer().element(xmlns, "oddHeader", *header_text);
		serializer().element(xmlns, "oddFooter", *footer_text);
		serializer().end_element(xmlns, "headerFooter");
	}

	if (!worksheet_rels.empty())
	{
		for (const auto &child_rel : worksheet_rels)
		{
			if (child_rel.get_type() == xlnt::relationship_type::vml_drawing)
			{
				serializer().start_element(xmlns, "legacyDrawing");
				serializer().attribute(xml::qname(xmlns_r, "id"), child_rel.get_id());
				serializer().end_element(xmlns, "legacyDrawing");

				//todo: there's only one of these per sheet, right?
				break;
			}
		}
	}
    
    serializer().end_element(xmlns, "worksheet");

	if (!worksheet_rels.empty())
	{
		write_relationships(worksheet_rels, worksheet_part);

		for (const auto &child_rel : worksheet_rels)
		{
			path archive_path(worksheet_part.parent().append(child_rel.get_target().get_path()));
			auto split_part_path = archive_path.split();
			auto part_path_iter = split_part_path.begin();
			while (part_path_iter != split_part_path.end())
			{
				if (*part_path_iter == "..")
				{
					part_path_iter = split_part_path.erase(part_path_iter - 1, part_path_iter + 1);
					continue;
				}

				++part_path_iter;
			}
			archive_path = std::accumulate(split_part_path.begin(), split_part_path.end(), path(""),
				[](const path &a, const std::string &b) { return a.append(b); });

            begin_part(archive_path);

			switch (child_rel.get_type())
			{
			case relationship::type::comments:
				write_comments(child_rel, ws, cells_with_comments);
				break;

			case relationship::type::vml_drawing:
				write_vml_drawings(child_rel, ws, cells_with_comments);
				break;

            case relationship::type::office_document: break;
            case relationship::type::thumbnail: break;
            case relationship::type::calculation_chain: break;
            case relationship::type::extended_properties: break;
            case relationship::type::core_properties: break;
            case relationship::type::worksheet: break;
            case relationship::type::shared_string_table: break;
            case relationship::type::styles: break;
            case relationship::type::theme: break;
            case relationship::type::hyperlink: break;
            case relationship::type::chartsheet: break;
            case relationship::type::unknown: break;
            case relationship::type::custom_properties: break;
            case relationship::type::printer_settings: break;
            case relationship::type::connections: break;
            case relationship::type::custom_property: break;
            case relationship::type::custom_xml_mappings: break;
            case relationship::type::dialogsheet: break;
            case relationship::type::drawings: break;
            case relationship::type::external_workbook_references: break;
            case relationship::type::metadata: break;
            case relationship::type::pivot_table: break;
            case relationship::type::pivot_table_cache_definition: break;
            case relationship::type::pivot_table_cache_records: break;
            case relationship::type::query_table: break;
            case relationship::type::shared_workbook_revision_headers: break;
            case relationship::type::shared_workbook: break;
            case relationship::type::revision_log: break;
            case relationship::type::shared_workbook_user_data: break;
            case relationship::type::single_cell_table_definitions: break;
            case relationship::type::table_definition: break;
            case relationship::type::volatile_dependencies: break;
            case relationship::type::image: break;
			}
		}
	}
}

// Sheet Relationship Target Parts

void xlsx_producer::write_comments(const relationship &/*rel*/, worksheet ws, 
	const std::vector<cell_reference> &cells)
{
	static const auto &xmlns = constants::get_namespace("worksheet");

	serializer().start_element(xmlns, "comments");
	serializer().namespace_decl(xmlns, "");

	if (!cells.empty())
	{
		std::unordered_map<std::string, std::size_t> authors;

		for (auto cell_ref : cells)
		{
			auto cell = ws.get_cell(cell_ref);
			auto author = cell.comment().author();

			if (authors.find(author) == authors.end())
			{
				authors[author] = authors.size();
			}
		}

		serializer().start_element(xmlns, "authors");

		for (const auto &author : authors)
		{
			serializer().start_element(xmlns, "author");
			serializer().characters(author.first);
			serializer().end_element(xmlns, "author");
		}

		serializer().end_element(xmlns, "authors");
		serializer().start_element(xmlns, "commentList");

		for (const auto &cell_ref : cells)
		{
			serializer().start_element(xmlns, "comment");

			auto cell = ws.get_cell(cell_ref);
			auto cell_comment = cell.comment();

			serializer().attribute("ref", cell_ref.to_string());
			auto author_id = authors.at(cell_comment.author());
			serializer().attribute("authorId", author_id);
			serializer().start_element(xmlns, "text");

			for (const auto &run : cell_comment.text().runs())
			{
				serializer().start_element(xmlns, "r");

				if (run.has_formatting())
				{
					serializer().start_element(xmlns, "rPr");

					if (run.has_size())
					{
						serializer().start_element(xmlns, "sz");
						serializer().attribute("val", run.get_size());
						serializer().end_element(xmlns, "sz");
					}

					if (run.has_color())
					{
						serializer().start_element(xmlns, "color");
						write_color(run.get_color());
						serializer().end_element(xmlns, "color");
					}

					if (run.has_font())
					{
						serializer().start_element(xmlns, "rFont");
						serializer().attribute("val", run.get_font());
						serializer().end_element(xmlns, "rFont");
					}

					if (run.has_family())
					{
						serializer().start_element(xmlns, "family");
						serializer().attribute("val", run.get_family());
						serializer().end_element(xmlns, "family");
					}

					if (run.has_scheme())
					{
						serializer().start_element(xmlns, "scheme");
						serializer().attribute("val", run.get_scheme());
						serializer().end_element(xmlns, "scheme");
					}

					if (run.bold_set())
					{
						serializer().start_element(xmlns, "b");
						serializer().end_element(xmlns, "b");
					}

					serializer().end_element(xmlns, "rPr");
				}

				serializer().element(xmlns, "t", run.get_string());
				serializer().end_element(xmlns, "r");
			}

			serializer().end_element(xmlns, "text");
			serializer().end_element(xmlns, "comment");
		}

		serializer().end_element(xmlns, "commentList");
	}

	serializer().end_element(xmlns, "comments");
}

void xlsx_producer::write_vml_drawings(const relationship &rel, worksheet ws,
    const std::vector<cell_reference> &cells)
{
    static const auto &xmlns_mv = std::string("http://macVmlSchemaUri");
    static const auto &xmlns_o = std::string("urn:schemas-microsoft-com:office:office");
    static const auto &xmlns_v = std::string("urn:schemas-microsoft-com:vml");
    static const auto &xmlns_x = std::string("urn:schemas-microsoft-com:office:excel");

    serializer().start_element("xml");
    serializer().namespace_decl(xmlns_v, "v");
    serializer().namespace_decl(xmlns_o, "o");
    serializer().namespace_decl(xmlns_x, "x");
    serializer().namespace_decl(xmlns_mv, "mv");
    
    serializer().start_element(xmlns_o, "shapelayout");
    serializer().attribute(xml::qname(xmlns_v, "ext"), "edit");
    serializer().start_element(xmlns_o, "idmap");
    serializer().attribute(xml::qname(xmlns_v, "ext"), "edit");

    auto filename = rel.get_target().get_path().split_extension().first;
    auto index_pos = filename.size() - 1;

    while (filename[index_pos] >= '0' && filename[index_pos] <= '9')
    {
	index_pos--;
    }

    auto file_index = std::stoull(filename.substr(index_pos + 1));

    serializer().attribute("data", file_index);
    serializer().end_element(xmlns_o, "idmap");
    serializer().end_element(xmlns_o, "shapelayout");

    serializer().start_element(xmlns_v, "shapetype");
    serializer().attribute("id", "_x0000_t202");
    serializer().attribute("coordsize", "21600,21600");
    serializer().attribute(xml::qname(xmlns_o, "spt"), "202");
    serializer().attribute("path", "m0,0l0,21600,21600,21600,21600,0xe");
    serializer().start_element(xmlns_v, "stroke");
    serializer().attribute("joinstyle", "miter");
    serializer().end_element(xmlns_v, "stroke");
    serializer().start_element(xmlns_v, "path");
    serializer().attribute("gradientshapeok", "t");
    serializer().attribute(xml::qname(xmlns_o, "connecttype"), "rect");
    serializer().end_element(xmlns_v, "path");
    serializer().end_element(xmlns_v, "shapetype");

    std::size_t comment_index = 0;

    for (const auto &cell_ref : cells)
    {
        auto comment = ws.get_cell(cell_ref).comment();
        auto shape_id = 1024 * file_index + 1 + comment_index * 2;

        serializer().start_element(xmlns_v, "shape");
        serializer().attribute("id", "_x0000_s" + std::to_string(shape_id));
        serializer().attribute("type", "#_x0000_t202");

        std::vector<std::pair<std::string, std::string>> style;

        style.push_back({"position", "absolute"});
        style.push_back({"margin-left", std::to_string(comment.left())});
        style.push_back({"margin-top", std::to_string(comment.top())});
        style.push_back({"width", std::to_string(comment.width())});
        style.push_back({"height", std::to_string(comment.height())});
        style.push_back({"z-index", std::to_string(comment_index + 1)});
        style.push_back({"visibility", comment.visible() ? "visible" : "hidden"});

        std::string style_string;

        for (auto part : style)
        {
            style_string.append(part.first);
            style_string.append(":");
            style_string.append(part.second);
            style_string.append(":");
        }

        serializer().attribute("style", style_string);
        serializer().attribute("fillcolor", "#fbf6d6");
        serializer().attribute("strokecolor", "#edeaa1");

        serializer().start_element(xmlns_v, "fill");
        serializer().attribute("color2", "#fbfe82");
        serializer().attribute("angle", -180);
        serializer().attribute("type", "gradient");
        serializer().start_element(xmlns_o, "fill");
        serializer().attribute(xml::qname(xmlns_v, "ext"), "view");
        serializer().attribute("type", "gradientUnscaled");
        serializer().end_element(xmlns_o, "fill");
        serializer().end_element(xmlns_v, "fill");

        serializer().start_element(xmlns_v, "shadow");
        serializer().attribute("on", "t");
        serializer().attribute("obscured", "t");
        serializer().end_element(xmlns_v, "shadow");

        serializer().start_element(xmlns_v, "path");
        serializer().attribute(xml::qname(xmlns_o, "connecttype"), "none");
        serializer().end_element(xmlns_v, "path");

        serializer().start_element(xmlns_v, "textbox");
        serializer().attribute("style", "mso-direction-alt:auto");
        serializer().start_element("div");
        serializer().attribute("style", "text-align:left");
        serializer().characters("");
        serializer().end_element("div");
        serializer().end_element(xmlns_v, "textbox");

        serializer().start_element(xmlns_x, "ClientData");
        serializer().attribute("ObjectType", "Note");
        serializer().start_element(xmlns_x, "MoveWithCells");
        serializer().end_element(xmlns_x, "MoveWithCells");
        serializer().start_element(xmlns_x, "SizeWithCells");
        serializer().end_element(xmlns_x, "SizeWithCells");
        serializer().start_element(xmlns_x, "Anchor");
        serializer().characters("1, 15, 0, " + std::to_string(2 + comment_index * 4) + ", 2, 54, 4, 14");
        serializer().end_element(xmlns_x, "Anchor");
        serializer().start_element(xmlns_x, "AutoFill");
        serializer().characters("False");
        serializer().end_element(xmlns_x, "AutoFill");
        serializer().start_element(xmlns_x, "Row");
        serializer().characters(cell_ref.get_row() - 1);
        serializer().end_element(xmlns_x, "Row");
        serializer().start_element(xmlns_x, "Column");
        serializer().characters(cell_ref.get_column_index().index - 1);
        serializer().end_element(xmlns_x, "Column");
        serializer().end_element(xmlns_x, "ClientData");

        serializer().end_element(xmlns_v, "shape");

        ++comment_index;
    }

    serializer().end_element("xml");
}

// Other Parts

void xlsx_producer::write_custom_property()
{
}

void xlsx_producer::write_unknown_parts()
{
}

void xlsx_producer::write_unknown_relationships()
{
}

void xlsx_producer::write_thumbnail(const relationship &rel)
{
    end_part();
    
    const auto &thumbnail = source_.get_thumbnail();
    vector_istreambuf thumbnail_buffer(thumbnail);
    archive_->open(rel.get_target().get_path().string()) << &thumbnail_buffer;
}

xml::serializer &xlsx_producer::serializer()
{
    return *current_part_serializer_;
}

std::string xlsx_producer::write_bool(bool boolean) const
{
	if (source_.d_->short_bools_)
	{
		return boolean ? "1" : "0";
	}

	return boolean ? "true" : "false";
}


void xlsx_producer::write_relationships(const std::vector<xlnt::relationship> &relationships, const path &part)
{
    path parent = part.parent();

    if (parent.is_absolute())
    {
        parent = path(parent.string().substr(1));
    }

	path rels_path(parent.append("_rels").append(part.filename() + ".rels").string());
    begin_part(rels_path);

    const auto xmlns = xlnt::constants::get_namespace("relationships");

    serializer().start_element(xmlns, "Relationships");
    serializer().namespace_decl(xmlns, "");

	for (const auto &relationship : relationships)
	{
        serializer().start_element(xmlns, "Relationship");

        serializer().attribute("Id", relationship.get_id());
        serializer().attribute("Type", relationship.get_type());
        serializer().attribute("Target", relationship.get_target().get_path().string());

		if (relationship.get_target_mode() == xlnt::target_mode::external)
		{
            serializer().attribute("TargetMode", "External");
		}

        serializer().end_element(xmlns, "Relationship");
	}
    
    serializer().end_element(xmlns, "Relationships");
}


void xlsx_producer::write_color(const xlnt::color &color)
{
	switch (color.get_type())
	{
	case xlnt::color::type::theme:
        serializer().attribute("theme", color.get_theme().get_index());
		break;

	case xlnt::color::type::indexed:
        serializer().attribute("indexed", color.get_indexed().get_index());
		break;

	case xlnt::color::type::auto_:
        serializer().attribute("auto", color.get_indexed().get_index());
		break;

	case xlnt::color::type::rgb:
        serializer().attribute("rgb", color.get_rgb().get_hex_string());
		break;
	}
}

} // namespace detail
} // namepsace xlnt
