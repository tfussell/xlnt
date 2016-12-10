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
#include <xlnt/worksheet/header_footer.hpp>
#include <xlnt/worksheet/worksheet.hpp>

using namespace std::string_literals;

namespace {

/// <summary>
/// Returns true if d is exactly equal to an integer.
/// </summary>
bool is_integral(long double d)
{
	return std::fabs(d - static_cast<long double>(static_cast<long long int>(d))) == 0.L;
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
    
    const auto root_rels = source_.manifest().relationships(path("/"));
    write_relationships(root_rels, path("/"));

	for (auto &rel : root_rels)
	{
        // thumbnail is binary content so we don't want to open an xml serializer stream
        if (rel.type() == relationship_type::thumbnail)
        {
            write_image(rel.target().path());
            continue;
        }

        begin_part(rel.target().path());

		switch (rel.type())
		{
		case relationship_type::core_properties:
			write_core_properties(rel);
			break;
            
		case relationship_type::extended_properties:
			write_extended_properties(rel);
			break;
            
		case relationship_type::custom_properties:
			write_custom_properties(rel);
			break;
            
		case relationship_type::office_document:
			write_workbook(rel);
			break;
        
        case relationship_type::thumbnail: break;
        case relationship_type::calculation_chain: break;
        case relationship_type::worksheet: break;
        case relationship_type::shared_string_table: break;
        case relationship_type::stylesheet: break;
        case relationship_type::theme: break;
        case relationship_type::hyperlink: break;
        case relationship_type::chartsheet: break;
        case relationship_type::comments: break;
        case relationship_type::vml_drawing: break;
        case relationship_type::unknown: break;
        case relationship_type::printer_settings: break;
        case relationship_type::connections: break;
        case relationship_type::custom_property: break;
        case relationship_type::custom_xml_mappings: break;
        case relationship_type::dialogsheet: break;
        case relationship_type::drawings: break;
        case relationship_type::external_workbook_references: break;
        case relationship_type::metadata: break;
        case relationship_type::pivot_table: break;
        case relationship_type::pivot_table_cache_definition: break;
        case relationship_type::pivot_table_cache_records: break;
        case relationship_type::query_table: break;
        case relationship_type::shared_workbook_revision_headers: break;
        case relationship_type::shared_workbook: break;
        case relationship_type::revision_log: break;
        case relationship_type::shared_workbook_user_data: break;
        case relationship_type::single_cell_table_definitions: break;
        case relationship_type::table_definition: break;
        case relationship_type::volatile_dependencies: break;
        case relationship_type::image: break;
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

	for (const auto &extension : source_.manifest().extensions_with_default_types())
	{
		serializer().start_element(xmlns, "Default");
		serializer().attribute("Extension", extension);
		serializer().attribute("ContentType",
            source_.manifest().default_type(extension));
		serializer().end_element(xmlns, "Default");
	}

	for (const auto &part : source_.manifest().parts_with_overriden_types())
	{
		serializer().start_element(xmlns, "Override");
		serializer().attribute("PartName", part.resolve(path("/")).string());
		serializer().attribute("ContentType",
            source_.manifest().override_type(part));
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

    if (source_.has_core_property("Application"))
    {
        serializer().element(xmlns, "Application", source_.core_property("Application"));
    }

    if (source_.has_core_property("DocSecurity"))
    {
        serializer().element(xmlns, "DocSecurity", source_.core_property("DocSecurity"));
    }

    if (source_.has_core_property("ScaleCrop"))
    {
        serializer().element(xmlns, "ScaleCrop", source_.core_property("ScaleCrop"));
    }

    serializer().start_element(xmlns, "HeadingPairs");
	serializer().start_element(xmlns_vt, "vector");
	serializer().attribute("size", "2");
	serializer().attribute("baseType", "variant");
    serializer().start_element(xmlns_vt, "variant");
    serializer().element(xmlns_vt, "lpstr", "Worksheets");
    serializer().end_element(xmlns_vt, "variant");
    serializer().start_element(xmlns_vt, "variant");
    serializer().element(xmlns_vt, "i4", source_.sheet_titles().size());
    serializer().end_element(xmlns_vt, "variant");
    serializer().end_element(xmlns_vt, "vector");
    serializer().end_element(xmlns, "HeadingPairs");

    serializer().start_element(xmlns, "TitlesOfParts");
    serializer().start_element(xmlns_vt, "vector");
    serializer().attribute("size", source_.sheet_titles().size());
    serializer().attribute("baseType", "lpstr");

	for (auto ws : source_)
	{
        serializer().element(xmlns_vt, "lpstr", ws.title());
	}

    serializer().end_element(xmlns_vt, "vector");
    serializer().end_element(xmlns, "TitlesOfParts");

	if (source_.has_core_property("Company"))
	{
        serializer().element(xmlns, "Company", source_.core_property("Company"));
	}

    if (source_.has_core_property("LinksUpToDate"))
    {
        serializer().element(xmlns, "LinksUpToDate", source_.core_property("LinksUpToDate"));
    }

    if (source_.has_core_property("LinksUpToDate"))
    {
        serializer().element(xmlns, "SharedDoc", source_.core_property("SharedDoc"));
    }
    
    if (source_.has_core_property("LinksUpToDate"))
    {
        serializer().element(xmlns, "HyperlinksChanged", source_.core_property("HyperlinksChanged"));
    }

    if (source_.has_core_property("LinksUpToDate"))
    {
        serializer().element(xmlns, "AppVersion", source_.core_property("AppVersion"));
    }
    
    serializer().end_element(xmlns, "Properties");
}

void xlsx_producer::write_core_properties(const relationship &/*rel*/)
{
    static const auto &xmlns_cp = constants::namespace_("core-properties");
    static const auto &xmlns_dc = constants::namespace_("dc");
    static const auto &xmlns_dcterms = constants::namespace_("dcterms");
    static const auto &xmlns_dcmitype = constants::namespace_("dcmitype");
    static const auto &xmlns_xsi = constants::namespace_("xsi");
    
	serializer().start_element(xmlns_cp, "coreProperties");
    serializer().namespace_decl(xmlns_cp, "cp");
    serializer().namespace_decl(xmlns_dc, "dc");
    serializer().namespace_decl(xmlns_dcterms, "dcterms");
    serializer().namespace_decl(xmlns_dcmitype, "dcmitype");
    serializer().namespace_decl(xmlns_xsi, "xsi");

    if (source_.has_core_property("creator"))
    {
        serializer().element(xmlns_dc, "creator", source_.core_property("creator"));
    }

    if (source_.has_core_property("lastModifiedBy"))
    {
        serializer().element(xmlns_cp, "lastModifiedBy", source_.core_property("lastModifiedBy"));
    }

    if (source_.has_core_property("created"))
    {
        serializer().start_element(xmlns_dcterms, "created");
        serializer().attribute(xmlns_xsi, "type", "dcterms:W3CDTF");
        serializer().characters(source_.core_property("created"));
        serializer().end_element(xmlns_dcterms, "created");
    }

    if (source_.has_core_property("modified"))
    {
        serializer().start_element(xmlns_dcterms, "modified");
        serializer().attribute(xmlns_xsi, "type", "dcterms:W3CDTF");
        serializer().characters(source_.core_property("modified"));
        serializer().end_element(xmlns_dcterms, "modified");
    }

	if (source_.has_title())
	{
		serializer().element(xmlns_dc, "title", source_.title());
	}

    if (source_.has_core_property("description"))
    {
        serializer().element(xmlns_dc, "description", source_.core_property("description"));
    }

    if (source_.has_core_property("subject"))
    {
        serializer().element(xmlns_dc, "subject", source_.core_property("subject"));
    }

    if (source_.has_core_property("keywords"))
    {
        serializer().element(xmlns_cp, "keywords", source_.core_property("keywords"));
    }

    if (source_.has_core_property("category"))
    {
        serializer().element(xmlns_cp, "category", source_.core_property("category"));
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
		if (!ws.has_page_setup() || ws.page_setup().sheet_state() == sheet_state::visible)
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

    static const auto &xmlns = constants::namespace_("workbook");
    static const auto &xmlns_r = constants::namespace_("r");
    static const auto &xmlns_s = constants::namespace_("spreadsheetml");

	serializer().start_element(xmlns, "workbook");
    serializer().namespace_decl(xmlns, "");
    serializer().namespace_decl(xmlns_r, "r");

	if (source_.has_file_version())
	{
		serializer().start_element(xmlns, "fileVersion");

		serializer().attribute("appName", source_.app_name());
		serializer().attribute("lastEdited", source_.last_edited());
		serializer().attribute("lowestEdited", source_.lowest_edited());
		serializer().attribute("rupBuild", source_.rup_build());
        
		serializer().end_element(xmlns, "fileVersion");
	}

	if (source_.has_code_name() || source_.base_date() == calendar::mac_1904)
	{
		serializer().start_element(xmlns, "workbookPr");

		if (source_.has_code_name())
		{
			serializer().attribute("codeName", source_.code_name());
		}

        if (source_.base_date() == calendar::mac_1904)
        {
            serializer().attribute("date1904", "1");
        }
        
		serializer().end_element(xmlns, "workbookPr");
	}

	if (source_.has_view())
	{
		serializer().start_element(xmlns, "bookViews");
		serializer().start_element(xmlns, "workbookView");

		const auto &view = source_.view();

        serializer().attribute("activeTab", "0");
        serializer().attribute("autoFilterDateGrouping", "1");
        serializer().attribute("firstSheet", "0");
        serializer().attribute("minimized", "0");
        serializer().attribute("showHorizontalScroll", "1");
        serializer().attribute("showSheetTabs", "1");
        serializer().attribute("showVerticalScroll", "1");
        serializer().attribute("visibility", "visible");

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
		auto sheet_rel_id = source_.d_->sheet_title_rel_id_map_[ws.title()];
		auto sheet_rel = source_.d_->manifest_.relationship(rel.target().path(), sheet_rel_id);

		serializer().start_element(xmlns, "sheet");
		serializer().attribute("name", ws.title());
		serializer().attribute("sheetId", ws.id());

		if (ws.has_page_setup() && ws.sheet_state() == xlnt::sheet_state::hidden)
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
			serializer().characters("'" + ws.title() + "'!"
                + range_reference::make_absolute(ws.auto_filter()).to_string());
            
			serializer().end_element(xmlns, "definedName");
		}
        
        serializer().end_element(xmlns, "sheet");
	}
#pragma clang diagnostic pop
    
    serializer().end_element(xmlns, "sheets");

    serializer().start_element(xmlns, "calcPr");
    serializer().attribute("calcId", 150000);
    serializer().attribute("calcMode", "auto");
    serializer().attribute("fullCalcOnLoad", "1");
    serializer().attribute("concurrentCalc", "0");
    serializer().end_element(xmlns, "calcPr");

	if (!source_.named_ranges().empty())
	{
        serializer().start_element(xmlns, "definedNames");

		for (auto &named_range : source_.named_ranges())
		{
            serializer().start_element(xmlns_s, "definedName");
            serializer().namespace_decl(xmlns_s, "s");
            serializer().attribute("name", named_range.name());
			const auto &target = named_range.targets().front();
			serializer().characters("'" + target.first.title()
                + "\'!" + target.second.to_string());
            serializer().end_element(xmlns_s, "definedName");
		}
        
        serializer().end_element(xmlns, "definedNames");
	}

    serializer().end_element(xmlns, "workbook");
    
    auto workbook_rels = source_.manifest().relationships(rel.target().path());
    write_relationships(workbook_rels, rel.target().path());
    
    for (const auto &child_rel : workbook_rels)
    {
		path archive_path(child_rel.source().path().parent().append(child_rel.target().path()));
        begin_part(archive_path);
        
        switch (child_rel.type())
        {
        case relationship_type::calculation_chain:
			write_calculation_chain(child_rel);
			break;
            
		case relationship_type::chartsheet:
			write_chartsheet(child_rel);
			break;
            
		case relationship_type::connections:
			write_connections(child_rel);
			break;
            
		case relationship_type::custom_xml_mappings:
			write_custom_xml_mappings(child_rel);
			break;
            
		case relationship_type::dialogsheet:
			write_dialogsheet(child_rel);
			break;
            
		case relationship_type::external_workbook_references:
			write_external_workbook_references(child_rel);
			break;
            
		case relationship_type::metadata:
			write_metadata(child_rel);
			break;
            
		case relationship_type::pivot_table:
			write_pivot_table(child_rel);
			break;

		case relationship_type::shared_string_table:
			write_shared_string_table(child_rel);
			break;

		case relationship_type::shared_workbook_revision_headers:
			write_shared_workbook_revision_headers(child_rel);
			break;
            
		case relationship_type::stylesheet:
			write_styles(child_rel);
			break;
            
		case relationship_type::theme:
			write_theme(child_rel);
			break;
            
		case relationship_type::volatile_dependencies:
			write_volatile_dependencies(child_rel);
			break;
            
		case relationship_type::worksheet:
			write_worksheet(child_rel);
			break;
            
        case relationship_type::office_document: break;
        case relationship_type::thumbnail: break;
        case relationship_type::extended_properties: break;
        case relationship_type::core_properties: break;
        case relationship_type::hyperlink: break;
        case relationship_type::comments: break;
        case relationship_type::vml_drawing: break;
        case relationship_type::unknown: break;
        case relationship_type::custom_properties: break;
        case relationship_type::printer_settings: break;
        case relationship_type::custom_property: break;
        case relationship_type::drawings: break;
        case relationship_type::pivot_table_cache_definition: break;
        case relationship_type::pivot_table_cache_records: break;
        case relationship_type::query_table: break;
        case relationship_type::shared_workbook: break;
        case relationship_type::revision_log: break;
        case relationship_type::shared_workbook_user_data: break;
        case relationship_type::single_cell_table_definitions: break;
        case relationship_type::table_definition: break;
        case relationship_type::image: break;
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
    static const auto &xmlns = constants::namespace_("spreadsheetml");

	serializer().start_element(xmlns, "sst");
    serializer().namespace_decl(xmlns, "");

    // todo: is there a more elegant way to get this number?
    std::size_t string_count = 0;

#pragma clang diagnostic push
#pragma clang diagnostic ignored "-Wrange-loop-analysis"
    for (const auto ws : source_)
    {
        auto dimension = ws.calculate_dimension();

        for (xlnt::row_t row = dimension.top_left().row();
            row <= dimension.bottom_right().row(); ++row)
        {
            for (xlnt::column_t column = dimension.top_left().column();
                column <= dimension.bottom_right().column(); ++column)
            {
                if (ws.has_cell(xlnt::cell_reference(column, row)))
                {
                    string_count += (ws.cell(xlnt::cell_reference(column, row))
                        .data_type() == cell::type::string) ? 1 : 0;
                }
            }
        }
    }
#pragma clang diagnostic pop

	serializer().attribute("count", string_count);
	serializer().attribute("uniqueCount", source_.shared_strings().size());

	for (const auto &string : source_.shared_strings())
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
                    serializer().attribute("val", run.size());
                    serializer().end_element(xmlns, "sz");
                }

                if (run.has_color())
                {
                    serializer().start_element(xmlns, "color");
                    write_color(run.color());
                    serializer().end_element(xmlns, "color");
                }

                if (run.has_font())
                {
                    serializer().start_element(xmlns, "rFont");
                    serializer().attribute("val", run.font());
                    serializer().end_element(xmlns, "rFont");
                }

                if (run.has_family())
                {
                    serializer().start_element(xmlns, "family");
                    serializer().attribute("val", run.family());
                    serializer().end_element(xmlns, "family");
                }

                if (run.has_scheme())
                {
                    serializer().start_element(xmlns, "scheme");
                    serializer().attribute("val", run.scheme());
                    serializer().end_element(xmlns, "scheme");
                }

				if (run.bold())
				{
					serializer().start_element(xmlns, "b");
					serializer().end_element(xmlns, "b");
				}
                
                serializer().end_element(xmlns, "rPr");
            }

            serializer().element(xmlns, "t", run.string());
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
    static const auto &xmlns = constants::namespace_("spreadsheetml");
    
	serializer().start_element(xmlns, "styleSheet");
	serializer().namespace_decl(xmlns, "");

	const auto &stylesheet = source_.impl().stylesheet_.get();

	// Number Formats

	if (!stylesheet.number_formats.empty())
	{
        const auto &number_formats = stylesheet.number_formats;

        auto num_custom = std::count_if(number_formats.begin(), number_formats.end(),
            [](const number_format &nf) { return nf.id() >= 164; });

		serializer().start_element(xmlns, "numFmts");
		serializer().attribute("count", num_custom);

		for (const auto &num_fmt : number_formats)
		{
            if (num_fmt.id() < 164) continue;
			serializer().start_element(xmlns, "numFmt");
			serializer().attribute("numFmtId", num_fmt.id());
			serializer().attribute("formatCode", num_fmt.format_string());
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

			if (current_font.has_color())
			{
                serializer().start_element(xmlns, "color");
                write_color(current_font.color());
                serializer().end_element(xmlns, "color");
			}

			serializer().start_element(xmlns, "name");
            serializer().attribute("val", current_font.name());
            serializer().end_element(xmlns, "name");

			if (current_font.has_family())
			{
                serializer().start_element(xmlns, "family");
                serializer().attribute("val", current_font.family());
                serializer().end_element(xmlns, "family");
			}

			if (current_font.has_scheme())
			{
                serializer().start_element(xmlns, "scheme");
                serializer().attribute("val", current_font.scheme());
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
        serializer().start_element(xmlns, "colors");
        serializer().start_element(xmlns, "indexedColors");

        for (auto &c : stylesheet.colors)
        {
            serializer().start_element(xmlns, "rgbColor");
            serializer().attribute("rgb", c.rgb().hex_string());
            serializer().end_element(xmlns, "rgbColor");
        }

        serializer().end_element(xmlns, "indexedColors");
        serializer().end_element(xmlns, "colors");
    }
    
    serializer().end_element(xmlns, "styleSheet");
}

void xlsx_producer::write_theme(const relationship &theme_rel)
{
    static const auto &xmlns_a = constants::namespace_("drawingml");
    static const auto &xmlns_thm15 = constants::namespace_("thm15");
    
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

    const auto workbook_rel = source_.manifest().relationship(path("/"), relationship_type::office_document);
    const auto theme_part = source_.manifest().canonicalize({workbook_rel, theme_rel});
    const auto theme_rels = source_.manifest().relationships(theme_part);

    if (!theme_rels.empty())
    {
        write_relationships(theme_rels, theme_part);

        for (auto rel : theme_rels)
        {
            if (rel.type() == relationship_type::image)
            {
                const auto image_path = source_.manifest().canonicalize({workbook_rel, theme_rel, rel});
                write_image(image_path);
            }
        }
    }
}

void xlsx_producer::write_volatile_dependencies(const relationship &/*rel*/)
{
	serializer().start_element("volTypes");
}

void xlsx_producer::write_worksheet(const relationship &rel)
{
	static const auto &xmlns = constants::namespace_("spreadsheetml");
	static const auto &xmlns_r = constants::namespace_("r");

	auto worksheet_part = rel.source().path().parent().append(rel.target().path());
	auto worksheet_rels = source_.manifest().relationships(worksheet_part);

	auto title = std::find_if(source_.d_->sheet_title_rel_id_map_.begin(),
		source_.d_->sheet_title_rel_id_map_.end(),
		[&](const std::pair<std::string, std::string> &p)
	{
		return p.second == rel.id();
	})->first;

	auto ws = source_.sheet_by_title(title);

	serializer().start_element(xmlns, "worksheet");
    serializer().namespace_decl(xmlns, "");
    serializer().namespace_decl(xmlns_r, "r");

	if (ws.has_page_setup())
	{
		serializer().start_element(xmlns, "sheetPr");

		serializer().start_element(xmlns, "outlinePr");
        serializer().attribute("summaryBelow", "1");
		serializer().attribute("summaryRight", "1");
        serializer().end_element(xmlns, "outlinePr");

		serializer().start_element(xmlns, "pageSetUpPr");
		serializer().attribute("fitToPage", write_bool(ws.page_setup().fit_to_page()));
		serializer().end_element(xmlns, "pageSetUpPr");

		serializer().end_element(xmlns, "sheetPr");
	}

    serializer().start_element(xmlns, "dimension");
    const auto dimension = ws.calculate_dimension();
    serializer().attribute("ref", dimension.is_single_cell()
        ? dimension.top_left().to_string() : dimension.to_string());
    serializer().end_element(xmlns, "dimension");

	if (ws.has_view())
	{
		serializer().start_element(xmlns, "sheetViews");
		serializer().start_element(xmlns, "sheetView");

		const auto view = ws.view();

		serializer().attribute("tabSelected", write_bool(view.id() == 0));
		serializer().attribute("workbookViewId", view.id());

		if (view.has_pane())
		{
            const auto &current_pane = view.pane();
            serializer().start_element(xmlns, "pane"); // CT_Pane

            if (current_pane.top_left_cell.is_set())
            {
                serializer().attribute("topLeftCell", current_pane.top_left_cell.get().to_string());
            }

            serializer().attribute("xSplit", current_pane.x_split.index);
            serializer().attribute("ySplit", current_pane.y_split);

            if (current_pane.active_pane != pane_corner::top_left)
            {
                serializer().attribute("activePane", current_pane.active_pane);
            }

            if (current_pane.state != pane_state::split)
            {
                serializer().attribute("state", current_pane.state);
            }

            serializer().end_element(xmlns, "pane");
		}

        for (const auto &current_selection : view.selections())
		{
			serializer().start_element(xmlns, "selection"); // CT_Selection

			if (current_selection.has_active_cell())
			{
				serializer().attribute("activeCell", current_selection.active_cell().to_string());
				serializer().attribute("sqref", current_selection.active_cell().to_string());
			}
/*
            if (current_selection.sqref() != "A1:A1")
            {
                serializer().attribute("sqref", current_selection.sqref().to_string());
            }
*/
            if (current_selection.pane() != pane_corner::top_left)
            {
                serializer().attribute("pane", current_selection.pane());
            }

            serializer().end_element(xmlns, "selection");
		}

		serializer().end_element(xmlns, "sheetView");
        serializer().end_element(xmlns, "sheetViews");
	}

    serializer().start_element(xmlns, "sheetFormatPr");
    serializer().attribute("baseColWidth", "10");
    serializer().attribute("defaultRowHeight", "16");
    serializer().end_element(xmlns, "sheetFormatPr");

	bool has_column_properties = false;

	for (auto column = ws.lowest_column(); column <= ws.highest_column(); column++)
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

		for (auto column = ws.lowest_column(); column <= ws.highest_column(); column++)
		{
			if (!ws.has_column_properties(column)) continue;

			const auto &props = ws.column_properties(column);

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

    const auto hyperlink_rels = source_.manifest().relationships(worksheet_part, relationship_type::hyperlink);
    std::unordered_map<std::string, std::string> reverse_hyperlink_references;
    
    for (auto hyperlink_rel : hyperlink_rels)
    {
        reverse_hyperlink_references[hyperlink_rel.target().path().string()] = rel.id();
    }

	std::unordered_map<std::string, std::string> hyperlink_references;

	serializer().start_element(xmlns, "sheetData");
	const auto &shared_strings = ws.workbook().shared_strings();
	std::vector<cell_reference> cells_with_comments;

	for (auto row : ws.rows())
	{
		auto min = static_cast<xlnt::row_t>(row.length());
		xlnt::row_t max = 0;
		bool any_non_null = false;

		for (auto cell : row)
		{
			min = std::min(min, cell.column().index);
			max = std::max(max, cell.column().index);

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

		serializer().attribute("r", row.front().row());
		serializer().attribute("spans", std::to_string(min) + ":" + std::to_string(max));

		if (ws.has_row_properties(row.front().row()))
		{
			serializer().attribute("customHeight", "1");
			auto height = ws.row_properties(row.front().row()).height;

			if (std::fabs(height - std::floor(height)) == 0.L)
			{
				serializer().attribute("ht", std::to_string(static_cast<long long int>(height)) + ".0");
			}
			else
			{
				serializer().attribute("ht", height);
			}
		}

		for (auto cell : row)
		{
			if (cell.has_comment())
			{
				cells_with_comments.push_back(cell.reference());
			}

			if (!cell.garbage_collectible())
			{
				serializer().start_element(xmlns, "c");
				serializer().attribute("r", cell.reference().to_string());
            
				if (cell.has_format())
				{
					serializer().attribute("s", cell.format().id());
				}

                if (cell.has_hyperlink())
                {
                    hyperlink_references[cell.reference().to_string()] = reverse_hyperlink_references[cell.hyperlink()];
                }

				if (cell.data_type() == cell::type::string)
				{
					if (cell.has_formula())
					{
						serializer().attribute("t", "str");
						serializer().element(xmlns, "f", cell.formula());
						serializer().element(xmlns, "v", cell.to_string());
                        serializer().end_element(xmlns, "c");

						continue;
					}

					int match_index = -1;

					for (std::size_t i = 0; i < shared_strings.size(); i++)
					{
						if (shared_strings[i] == cell.value<formatted_text>())
						{
							match_index = static_cast<int>(i);
							break;
						}
					}

					if (match_index == -1)
					{
						if (cell.value<std::string>().empty())
						{
							serializer().attribute("t", "s");
						}
						else
						{
							serializer().attribute("t", "inlineStr");
							serializer().start_element(xmlns, "is");
							serializer().element(xmlns, "t", cell.value<std::string>());
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
					if (cell.data_type() != cell::type::null)
					{
						if (cell.data_type() == cell::type::boolean)
						{
							serializer().attribute("t", "b");
							serializer().element(xmlns, "v", write_bool(cell.value<bool>()));
						}
						else if (cell.data_type() == cell::type::numeric)
						{
							if (cell.has_formula())
							{
								serializer().element(xmlns, "f", cell.formula());
								serializer().element(xmlns, "v", cell.to_string());
                                serializer().end_element(xmlns, "c");

								continue;
							}

							serializer().attribute("t", "n");
							serializer().start_element(xmlns, "v");

							if (is_integral(cell.value<long double>()))
							{
                                serializer().characters(cell.value<long long>());
							}
							else
							{
								std::stringstream ss;
								ss.precision(20);
								ss << cell.value<long double>();
								ss.str();
                                serializer().characters(ss.str());
							}
                            
                            serializer().end_element(xmlns, "v");
						}
					}
					else if (cell.has_formula())
					{
                        serializer().element(xmlns, "f", cell.formula());
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
		serializer().attribute("ref", ws.auto_filter().to_string());
		serializer().end_element(xmlns, "autoFilter");
	}

	if (!ws.merged_ranges().empty())
	{
		serializer().start_element(xmlns, "mergeCells");
		serializer().attribute("count", ws.merged_ranges().size());

		for (auto merged_range : ws.merged_ranges())
		{
            serializer().start_element(xmlns, "mergeCell");
			serializer().attribute("ref", merged_range.to_string());
            serializer().end_element(xmlns, "mergeCell");
		}
        
		serializer().end_element(xmlns, "mergeCells");
	}

	if (!hyperlink_rels.empty())
	{
        serializer().start_element(xmlns, "hyperlinks");

        for (const auto &hyperlink : hyperlink_references)
        {
            serializer().start_element(xmlns, "hyperlink");
            serializer().attribute("ref", hyperlink.first);
            serializer().attribute(xmlns_r, "id", hyperlink.second);
            serializer().end_element(xmlns, "hyperlink");
        }
        
        serializer().end_element(xmlns, "hyperlinks");
	}

	if (ws.has_page_setup())
	{
		serializer().start_element(xmlns, "printOptions");
		serializer().attribute("horizontalCentered", write_bool(ws.page_setup().horizontal_centered()));
		serializer().attribute("verticalCentered", write_bool(ws.page_setup().vertical_centered()));
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

		serializer().attribute("left", remove_trailing_zeros(std::to_string(ws.page_margins().left())));
		serializer().attribute("right", remove_trailing_zeros(std::to_string(ws.page_margins().right())));
		serializer().attribute("top", remove_trailing_zeros(std::to_string(ws.page_margins().top())));
		serializer().attribute("bottom", remove_trailing_zeros(std::to_string(ws.page_margins().bottom())));
		serializer().attribute("header", remove_trailing_zeros(std::to_string(ws.page_margins().header())));
		serializer().attribute("footer", remove_trailing_zeros(std::to_string(ws.page_margins().footer())));
        
        serializer().end_element(xmlns, "pageMargins");
	}

	if (ws.has_page_setup())
	{
		serializer().start_element(xmlns, "pageSetup");
		serializer().attribute("orientation",
            ws.page_setup().orientation() == xlnt::orientation::landscape
            ? "landscape" : "portrait");
		serializer().attribute("paperSize",  static_cast<std::size_t>(ws.page_setup().paper_size()));
        serializer().attribute("fitToHeight", write_bool(ws.page_setup().fit_to_height()));
		serializer().attribute("fitToWidth", write_bool(ws.page_setup().fit_to_width()));
		serializer().end_element(xmlns, "pageSetup");
	}

	if (ws.has_header_footer())
	{
        const auto hf = ws.header_footer();

		serializer().start_element(xmlns, "headerFooter");

        auto odd_header = std::string();
        auto odd_footer = std::string();
        auto even_header = std::string();
        auto even_footer = std::string();
        auto first_header = std::string();
        auto first_footer = std::string();

        const auto encode_text = [](const formatted_text &t, header_footer::location where)
        {
            const auto location_code_map = std::unordered_map<xlnt::header_footer::location, std::string>
            {
                { xlnt::header_footer::location::left, "&L" },
                { xlnt::header_footer::location::center, "&C" },
                { xlnt::header_footer::location::right, "&R" },
            };
            
            return location_code_map.at(where) + t.plain_text();
        };

        const auto locations =
        {
            header_footer::location::left,
            header_footer::location::center,
            header_footer::location::right
        };

        for (auto location : locations)
        {
            if (hf.different_odd_even())
            {
                if (hf.has_odd_even_header(location))
                {
                    odd_header.append(encode_text(hf.odd_header(location), location));
                    even_header.append(encode_text(hf.even_header(location), location));
                }

                if (hf.has_odd_even_footer(location))
                {
                    odd_footer.append(encode_text(hf.odd_footer(location), location));
                    even_footer.append(encode_text(hf.even_footer(location), location));
                }
            }
            else
            {
                if (hf.has_header(location))
                {
                    odd_header.append(encode_text(hf.header(location), location));
                }

                if (hf.has_footer(location))
                {
                    odd_footer.append(encode_text(hf.footer(location), location));
                }
            }

            if (hf.different_first())
            {
                if (hf.has_first_page_header(location))
                {
                    first_header.append(encode_text(hf.first_page_header(location), location));
                }

                if (hf.has_first_page_footer(location))
                {
                    first_footer.append(encode_text(hf.first_page_footer(location), location));
                }
            }
        }

        if (!odd_header.empty())
        {
            serializer().element(xmlns, "oddHeader", odd_header);
        }

        if (!odd_footer.empty())
        {
            serializer().element(xmlns, "oddFooter", odd_footer);
        }

        if (!even_header.empty())
        {
            serializer().element(xmlns, "evenHeader", even_header);
        }

        if (!even_footer.empty())
        {
            serializer().element(xmlns, "evenFooter", even_footer);
        }

        if (!first_header.empty())
        {
            serializer().element(xmlns, "firstHeader", first_header);
        }

        if (!first_footer.empty())
        {
            serializer().element(xmlns, "firstFooter", first_footer);
        }

		serializer().end_element(xmlns, "headerFooter");
	}

	if (!worksheet_rels.empty())
	{
		for (const auto &child_rel : worksheet_rels)
		{
			if (child_rel.type() == xlnt::relationship_type::vml_drawing)
			{
				serializer().start_element(xmlns, "legacyDrawing");
				serializer().attribute(xml::qname(xmlns_r, "id"), child_rel.id());
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
            if (child_rel.target_mode() == target_mode::external) continue;

            // todo: this is ugly
			path archive_path(worksheet_part.parent().append(child_rel.target().path()));
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

			switch (child_rel.type())
			{
			case relationship_type::comments:
				write_comments(child_rel, ws, cells_with_comments);
				break;

			case relationship_type::vml_drawing:
				write_vml_drawings(child_rel, ws, cells_with_comments);
				break;

            case relationship_type::office_document: break;
            case relationship_type::thumbnail: break;
            case relationship_type::calculation_chain: break;
            case relationship_type::extended_properties: break;
            case relationship_type::core_properties: break;
            case relationship_type::worksheet: break;
            case relationship_type::shared_string_table: break;
            case relationship_type::stylesheet: break;
            case relationship_type::theme: break;
            case relationship_type::hyperlink: break;
            case relationship_type::chartsheet: break;
            case relationship_type::unknown: break;
            case relationship_type::custom_properties: break;
            case relationship_type::printer_settings: break;
            case relationship_type::connections: break;
            case relationship_type::custom_property: break;
            case relationship_type::custom_xml_mappings: break;
            case relationship_type::dialogsheet: break;
            case relationship_type::drawings: break;
            case relationship_type::external_workbook_references: break;
            case relationship_type::metadata: break;
            case relationship_type::pivot_table: break;
            case relationship_type::pivot_table_cache_definition: break;
            case relationship_type::pivot_table_cache_records: break;
            case relationship_type::query_table: break;
            case relationship_type::shared_workbook_revision_headers: break;
            case relationship_type::shared_workbook: break;
            case relationship_type::revision_log: break;
            case relationship_type::shared_workbook_user_data: break;
            case relationship_type::single_cell_table_definitions: break;
            case relationship_type::table_definition: break;
            case relationship_type::volatile_dependencies: break;
            case relationship_type::image: break;
			}
		}
	}
}

// Sheet Relationship Target Parts

void xlsx_producer::write_comments(const relationship &/*rel*/, worksheet ws, 
	const std::vector<cell_reference> &cells)
{
	static const auto &xmlns = constants::namespace_("spreadsheetml");

	serializer().start_element(xmlns, "comments");
	serializer().namespace_decl(xmlns, "");

	if (!cells.empty())
	{
		std::unordered_map<std::string, std::size_t> authors;

		for (auto cell_ref : cells)
		{
			auto cell = ws.cell(cell_ref);
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

			auto cell = ws.cell(cell_ref);
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
						serializer().attribute("val", run.size());
						serializer().end_element(xmlns, "sz");
					}

					if (run.has_color())
					{
						serializer().start_element(xmlns, "color");
						write_color(run.color());
						serializer().end_element(xmlns, "color");
					}

					if (run.has_font())
					{
						serializer().start_element(xmlns, "rFont");
						serializer().attribute("val", run.font());
						serializer().end_element(xmlns, "rFont");
					}

					if (run.has_family())
					{
						serializer().start_element(xmlns, "family");
						serializer().attribute("val", run.family());
						serializer().end_element(xmlns, "family");
					}

					if (run.has_scheme())
					{
						serializer().start_element(xmlns, "scheme");
						serializer().attribute("val", run.scheme());
						serializer().end_element(xmlns, "scheme");
					}

					if (run.bold())
					{
						serializer().start_element(xmlns, "b");
						serializer().end_element(xmlns, "b");
					}

					serializer().end_element(xmlns, "rPr");
				}

				serializer().element(xmlns, "t", run.string());
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

    auto filename = rel.target().path().split_extension().first;
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
        auto comment = ws.cell(cell_ref).comment();
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
            style_string.append(";");
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
        serializer().characters(cell_ref.row() - 1);
        serializer().end_element(xmlns_x, "Row");
        serializer().start_element(xmlns_x, "Column");
        serializer().characters(cell_ref.column_index() - 1);
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

void xlsx_producer::write_image(const path &image_path)
{
    end_part();
    
    vector_istreambuf buffer(source_.d_->images_.at(image_path.string()));
    archive_->open(image_path.string()) << &buffer;
}

xml::serializer &xlsx_producer::serializer()
{
    return *current_part_serializer_;
}

std::string xlsx_producer::write_bool(bool boolean) const
{
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

    const auto xmlns = xlnt::constants::namespace_("relationships");

    serializer().start_element(xmlns, "Relationships");
    serializer().namespace_decl(xmlns, "");

	for (const auto &relationship : relationships)
	{
        serializer().start_element(xmlns, "Relationship");

        serializer().attribute("Id", relationship.id());
        serializer().attribute("Type", relationship.type());
        serializer().attribute("Target", relationship.target().path().string());

		if (relationship.target_mode() == xlnt::target_mode::external)
		{
            serializer().attribute("TargetMode", "External");
		}

        serializer().end_element(xmlns, "Relationship");
	}
    
    serializer().end_element(xmlns, "Relationships");
}


void xlsx_producer::write_color(const xlnt::color &color)
{
    if (color.is_auto())
    {
        serializer().attribute("auto", write_bool(true));
        return;
    }

	switch (color.type())
	{
	case xlnt::color_type::theme:
        serializer().attribute("theme", color.theme().index());
		break;

	case xlnt::color_type::indexed:
        serializer().attribute("indexed", color.indexed().index());
		break;

	case xlnt::color_type::rgb:
        serializer().attribute("rgb", color.rgb().hex_string());
		break;
	}
}

} // namespace detail
} // namepsace xlnt
