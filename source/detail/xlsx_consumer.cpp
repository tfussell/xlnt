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

#include <cctype>
#include <numeric> // for std::accumulate

#include <detail/xlsx_consumer.hpp>
#include <detail/constants.hpp>
#include <detail/custom_value_traits.hpp>
#include <detail/workbook_impl.hpp>
#include <detail/zip.hpp>
#include <xlnt/cell/cell.hpp>
#include <xlnt/utils/path.hpp>
#include <xlnt/packaging/manifest.hpp>
#include <xlnt/workbook/const_worksheet_iterator.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/worksheet/worksheet.hpp>

namespace {

bool is_true(const std::string &bool_string)
{
	return bool_string == "1" || bool_string == "true";
}

std::size_t string_to_size_t(const std::string &s)
{
#if ULLONG_MAX == SIZE_MAX
	return std::stoull(s);
#else
	return std::stoul(s);
#endif
}

xlnt::datetime w3cdtf_to_datetime(const std::string &string)
{
	xlnt::datetime result(1900, 1, 1);
	auto separator_index = string.find('-');
	result.year = std::stoi(string.substr(0, separator_index));
	result.month = std::stoi(string.substr(separator_index + 1, string.find('-', separator_index + 1)));
	separator_index = string.find('-', separator_index + 1);
	result.day = std::stoi(string.substr(separator_index + 1, string.find('T', separator_index + 1)));
	separator_index = string.find('T', separator_index + 1);
	result.hour = std::stoi(string.substr(separator_index + 1, string.find(':', separator_index + 1)));
	separator_index = string.find(':', separator_index + 1);
	result.minute = std::stoi(string.substr(separator_index + 1, string.find(':', separator_index + 1)));
	separator_index = string.find(':', separator_index + 1);
	result.second = std::stoi(string.substr(separator_index + 1, string.find('Z', separator_index + 1)));
	return result;
}

xlnt::color read_color(xml::parser &parser)
{
	xlnt::color result;

	if (parser.attribute_present("auto"))
	{
		return result;
	}

	if (parser.attribute_present("rgb"))
	{
		result = xlnt::rgb_color(parser.attribute("rgb"));
	}
	else if (parser.attribute_present("theme"))
	{
		result = xlnt::theme_color(string_to_size_t(parser.attribute("theme")));
	}
	else if (parser.attribute_present("indexed"))
	{
		result = xlnt::indexed_color(string_to_size_t(parser.attribute("indexed")));
	}

	if (parser.attribute_present("tint"))
	{
		result.set_tint(parser.attribute("tint", 0.0));
	}

	return result;
}

std::vector<xlnt::relationship> read_relationships(const xlnt::path &part, xlnt::detail::ZipFileReader &archive)
{
    std::vector<xlnt::relationship> relationships;
    if (!archive.has_file(part.string())) return relationships;

    auto &rels_stream = archive.open(part.string());
    xml::parser parser(rels_stream, part.string());

    xlnt::uri source(part.string());

    static const auto &xmlns = xlnt::constants::get_namespace("relationships");
    parser.next_expect(xml::parser::event_type::start_element, xmlns, "Relationships");
    parser.content(xml::content::complex);

    while (true)
	{
        if (parser.peek() == xml::parser::event_type::end_element) break;
        
        parser.next_expect(xml::parser::event_type::start_element, xmlns, "Relationship");
		relationships.emplace_back(parser.attribute("Id"),
            parser.attribute<xlnt::relationship::type>("Type"), source,
            xlnt::uri(parser.attribute("Target")), xlnt::target_mode::internal);
        parser.next_expect(xml::parser::event_type::end_element,
            xlnt::constants::get_namespace("relationships"), "Relationship");
	}
    
    parser.next_expect(xml::parser::event_type::end_element, xmlns, "Relationships");

	return relationships;
}

void check_document_type(const std::string &document_content_type)
{
	if (document_content_type != "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
		&& document_content_type != "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml ")
	{
		throw xlnt::invalid_file(document_content_type);
	}
}

} // namespace

namespace xlnt {
namespace detail {

xlsx_consumer::xlsx_consumer(workbook &target) 
	: target_(target),
	  parser_(nullptr)
{
}

void xlsx_consumer::read(std::istream &source)
{
	archive_.reset(new ZipFileReader(source));
	populate_workbook();
}

xml::parser &xlsx_consumer::parser()
{
	return *parser_;
}

void xlsx_consumer::populate_workbook()
{
	target_.clear();

	auto &manifest = target_.get_manifest();
	read_manifest();

	for (const auto &rel : manifest.get_relationships(path("/")))
	{
        xml::parser parser(archive_->open(rel.get_target().get_path().string()),
            rel.get_target().get_path().string());
		parser_ = &parser;

		switch (rel.get_type())
		{
		case relationship::type::core_properties:
			read_core_properties();
			break;

		case relationship::type::extended_properties:
			read_extended_properties();
			break;

		case relationship::type::custom_properties:
			read_custom_property();
			break;

		case relationship::type::office_document:
            check_document_type(manifest.get_content_type(rel.get_target().get_path()));
			read_workbook();
			break;

		case relationship::type::connections:
			read_connections();
			break;

		case relationship::type::custom_xml_mappings:
			read_custom_xml_mappings();
			break;

		case relationship::type::external_workbook_references:
			read_external_workbook_references();
			break;

		case relationship::type::metadata:
			read_metadata();
			break;

		case relationship::type::pivot_table:
			read_pivot_table();
			break;

		case relationship::type::shared_workbook_revision_headers:
			read_shared_workbook_revision_headers();
			break;

		case relationship::type::volatile_dependencies:
			read_volatile_dependencies();
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
        case relationship::type::custom_property: break;
        case relationship::type::dialogsheet: break;
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

		parser_ = nullptr;
	}

    const auto workbook_rel = manifest.get_relationship(path("/"), relationship::type::office_document);

    // First pass of workbook relationship parts which must be read before sheets (e.g. shared strings)
    
	for (const auto &rel : manifest.get_relationships(workbook_rel.get_target().get_path()))
	{
		path part_path(rel.get_source().get_path().parent().append(rel.get_target().get_path()));
        auto receive = xml::parser::receive_default;
        auto using_namespaces = rel.get_type() == relationship::type::styles;
	if (using_namespaces)
	{
            receive |= xml::parser::receive_namespace_decls;
	}
        xml::parser parser(archive_->open(part_path.string()), part_path.string(), receive);
		parser_ = &parser;

		switch (rel.get_type())
		{
        case relationship::type::shared_string_table:
            read_shared_string_table();
            break;

        case relationship::type::styles:
            read_stylesheet();
            break;

        case relationship::type::theme:
            read_theme();
            break;

        case relationship::type::office_document: break;
        case relationship::type::thumbnail: break;
        case relationship::type::calculation_chain: break;
        case relationship::type::extended_properties: break;
        case relationship::type::core_properties: break;
        case relationship::type::worksheet: break;
        case relationship::type::hyperlink: break;
        case relationship::type::chartsheet: break;
        case relationship::type::comments: break;
        case relationship::type::vml_drawing: break;
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

		parser_ = nullptr;
	}
    
    // Second pass, read sheets themselves

	for (const auto &rel : manifest.get_relationships(workbook_rel.get_target().get_path()))
	{
	    path part_path(rel.get_source().get_path().parent().append(rel.get_target().get_path()));
	    auto receive = xml::parser::receive_default;
	    receive |= xml::parser::receive_namespace_decls;
	    xml::parser parser(archive_->open(part_path.string()), rel.get_target().get_path().string(), receive);
	    parser_ = &parser;

		switch (rel.get_type())
		{
		case relationship::type::chartsheet:
			read_chartsheet(rel.get_id());
			break;

		case relationship::type::dialogsheet:
			read_dialogsheet(rel.get_id());
			break;

		case relationship::type::worksheet:
			read_worksheet(rel.get_id());
			break;

        case relationship::type::office_document: break;
        case relationship::type::thumbnail: break;
        case relationship::type::calculation_chain: break;
        case relationship::type::extended_properties: break;
        case relationship::type::core_properties: break;
        case relationship::type::shared_string_table: break;
        case relationship::type::styles: break;
        case relationship::type::theme: break;
        case relationship::type::hyperlink: break;
        case relationship::type::comments: break;
        case relationship::type::vml_drawing: break;
        case relationship::type::unknown: break;
        case relationship::type::custom_properties: break;
        case relationship::type::printer_settings: break;
        case relationship::type::connections: break;
        case relationship::type::custom_property: break;
        case relationship::type::custom_xml_mappings: break;
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

		parser_ = nullptr;
	}

	// Unknown Parts

	void read_unknown_parts();
	void read_unknown_relationships();
}

// Package Parts

void xlsx_consumer::read_manifest()
{
	path package_rels_path("_rels/.rels");

	if (!archive_->has_file(package_rels_path.string()))
	{
		throw invalid_file("missing package rels");
	}

	auto package_rels = read_relationships(package_rels_path, *archive_);
	auto &manifest = target_.get_manifest();

    static const auto &xmlns = constants::get_namespace("content-types");

    xml::parser parser(archive_->open("[Content_Types].xml"), "[Content_Types].xml");
    parser.next_expect(xml::parser::event_type::start_element, xmlns, "Types");
    parser.content(xml::content::complex);

    while (true)
	{
        if (parser.peek() == xml::parser::event_type::end_element) break;
        
        parser.next_expect(xml::parser::event_type::start_element);
        
        if (parser.name() == "Default")
        {
            manifest.register_default_type(parser.attribute("Extension"),
				parser.attribute("ContentType"));
            parser.next_expect(xml::parser::event_type::end_element, xmlns, "Default");
        }
        else if (parser.name() == "Override")
        {
			manifest.register_override_type(path(parser.attribute("PartName")),
                parser.attribute("ContentType"));
            parser.next_expect(xml::parser::event_type::end_element, xmlns, "Override");
        }
	}
    
    parser.next_expect(xml::parser::event_type::end_element, xmlns, "Types");

	for (const auto &package_rel : package_rels)
	{
		manifest.register_relationship(uri("/"),
            package_rel.get_type(),
			package_rel.get_target(),
			package_rel.get_target_mode(), 
			package_rel.get_id());
	}

	for (const auto &relationship_source_string : archive_->files())
	{
		auto relationship_source = path(relationship_source_string);

		if (relationship_source == path("_rels/.rels") 
			|| relationship_source.extension() != "rels") continue;

		path part(relationship_source.parent().parent());
		part = part.append(relationship_source.split_extension().first);
		uri source(part.string());

		path source_directory = part.parent();

		auto part_rels = read_relationships(relationship_source, *archive_);

		for (const auto &part_rel : part_rels)
		{
			path target_path(source_directory.append(part_rel.get_target().get_path()));
			manifest.register_relationship(source, part_rel.get_type(),
				part_rel.get_target(), part_rel.get_target_mode(), part_rel.get_id());
		}
	}
}

void xlsx_consumer::read_extended_properties()
{
    static const auto &xmlns = constants::get_namespace("extended-properties");
    static const auto &xmlns_vt = constants::get_namespace("vt");

    parser().next_expect(xml::parser::event_type::start_element, xmlns, "Properties");
    parser().content(xml::parser::content_type::complex);
    
    while (true)
	{
        if (parser().peek() == xml::parser::event_type::end_element) break;
        
        parser().next_expect(xml::parser::event_type::start_element);

        auto name = parser().name();
        auto text = std::string();

        while (parser().peek() == xml::parser::event_type::characters)
        {
            parser().next_expect(xml::parser::event_type::characters);
            text.append(parser().value());
        }
        
        if (name == "Application") target_.set_application(text);
        else if (name == "DocSecurity") target_.set_doc_security(std::stoi(text));
        else if (name == "ScaleCrop") target_.set_scale_crop(is_true(text));
        else if (name == "Company") target_.set_company(text);
        else if (name == "SharedDoc") target_.set_shared_doc(is_true(text));
        else if (name == "HyperlinksChanged") target_.set_hyperlinks_changed(is_true(text));
        else if (name == "AppVersion") target_.set_app_version(text);
        else if (name == "Application") target_.set_application(text);
        else if (name == "HeadingPairs")
        {
            parser().next_expect(xml::parser::event_type::start_element, xmlns_vt, "vector");
            parser().content(xml::parser::content_type::complex);

            parser().attribute("size");
            parser().attribute("baseType");

            parser().next_expect(xml::parser::event_type::start_element, xmlns_vt, "variant");
            parser().content(xml::parser::content_type::complex);
            parser().next_expect(xml::parser::event_type::start_element, xmlns_vt, "lpstr");
            parser().next_expect(xml::parser::event_type::characters);
            parser().next_expect(xml::parser::event_type::end_element, xmlns_vt, "lpstr");
            parser().next_expect(xml::parser::event_type::end_element, xmlns_vt, "variant");
            parser().next_expect(xml::parser::event_type::start_element, xmlns_vt, "variant");
            parser().content(xml::parser::content_type::complex);
            parser().next_expect(xml::parser::event_type::start_element, xmlns_vt, "i4");
            parser().next_expect(xml::parser::event_type::characters);
            parser().next_expect(xml::parser::event_type::end_element, xmlns_vt, "i4");
            parser().next_expect(xml::parser::event_type::end_element, xmlns_vt, "variant");
            
            parser().next_expect(xml::parser::event_type::end_element, xmlns_vt, "vector");
        }
        else if (name == "TitlesOfParts")
        {
            parser().next_expect(xml::parser::event_type::start_element, xmlns_vt, "vector");
            parser().content(xml::parser::content_type::complex);

            parser().attribute("size");
            parser().attribute("baseType");
            
            while (true)
            {
                if (parser().peek() == xml::parser::event_type::end_element) break;

                parser().next_expect(xml::parser::event_type::start_element, xmlns_vt, "lpstr");
                parser().content(xml::parser::content_type::simple);
                parser().next_expect(xml::parser::event_type::characters);
                parser().next_expect(xml::parser::event_type::end_element, xmlns_vt, "lpstr");
            }
            
            parser().next_expect(xml::parser::event_type::end_element, xmlns_vt, "vector");
        }
        
        while (parser().peek() == xml::parser::event_type::characters)
        {
            parser().next_expect(xml::parser::event_type::characters);
        }
        
        parser().next_expect(xml::parser::event_type::end_element);
	}
}

void xlsx_consumer::read_core_properties()
{
    static const auto &xmlns_cp = constants::get_namespace("core-properties");
    static const auto &xmlns_dc = constants::get_namespace("dc");
    static const auto &xmlns_dcterms = constants::get_namespace("dcterms");
    static const auto &xmlns_xsi = constants::get_namespace("xsi");

    parser().next_expect(xml::parser::event_type::start_element, xmlns_cp, "coreProperties");
    parser().content(xml::parser::content_type::complex);
    
    while (true)
	{
        if (parser().peek() == xml::parser::event_type::end_element) break;
        
        parser().next_expect(xml::parser::event_type::start_element);

        std::string characters;
        if (parser().peek() == xml::parser::event_type::characters)
        {
            parser().next_expect(xml::parser::event_type::characters);
            characters = parser().value();
        }

		if (parser().namespace_() == xmlns_dc && parser().name() == "creator")
        {
            target_.set_creator(characters);
        }
		else if (parser().namespace_() == xmlns_cp && parser().name() == "lastModifiedBy")
        {
            target_.set_last_modified_by(characters);
        }
		else if (parser().namespace_() == xmlns_dcterms && parser().name() == "created")
        {
            parser().attribute(xml::qname(xmlns_xsi, "type"));
            target_.set_created(w3cdtf_to_datetime(characters));
        }
		else if (parser().namespace_() == xmlns_dcterms && parser().name() == "modified")
        {
            parser().attribute(xml::qname(xmlns_xsi, "type"));
            target_.set_modified(w3cdtf_to_datetime(characters));
        }
        
        parser().next_expect(xml::parser::event_type::end_element);
	}
    
    parser().next_expect(xml::parser::event_type::end_element, xmlns_cp, "coreProperties");
}

void xlsx_consumer::read_custom_file_properties()
{
}

// Write SpreadsheetML-Specific Package Parts

void xlsx_consumer::read_workbook()
{
    static const auto &xmlns = constants::get_namespace("workbook");
    static const auto &xmlns_mc = constants::get_namespace("mc");
    static const auto &xmlns_mx = constants::get_namespace("mx");
    static const auto &xmlns_r = constants::get_namespace("r");
    static const auto &xmlns_s = constants::get_namespace("worksheet");
    static const auto &xmlns_x15 = constants::get_namespace("x15");
    static const auto &xmlns_x15ac = constants::get_namespace("x15ac");

    parser().next_expect(xml::parser::event_type::start_element, xmlns, "workbook");
    parser().content(xml::parser::content_type::complex);
    
	while (parser().peek() == xml::parser::event_type::start_namespace_decl)
	{
        parser().next_expect(xml::parser::event_type::start_namespace_decl);
        if (parser().name() == "x15") target_.enable_x15();
        parser().next_expect(xml::parser::event_type::end_namespace_decl);
	}
    
    if (parser().attribute_present(xml::qname(xmlns_mc, "Ignorable")))
    {
        parser().attribute(xml::qname(xmlns_mc, "Ignorable"));
    }

    while (true)
    {
        if (parser().peek() == xml::parser::event_type::end_element) break;
        
        parser().next_expect(xml::parser::event_type::start_element);
        parser().content(xml::parser::content_type::complex);

        auto qname = parser().qname();

        if (qname == xml::qname(xmlns, "fileVersion"))
        {
            target_.d_->has_file_version_ = true;
            target_.d_->file_version_.app_name = parser().attribute("appName");
            if (parser().attribute_present("lastEdited"))
            {
                target_.d_->file_version_.last_edited = string_to_size_t(parser().attribute("lastEdited"));
            }
            if (parser().attribute_present("lowestEdited"))
            {
                target_.d_->file_version_.lowest_edited = string_to_size_t(parser().attribute("lowestEdited"));
            }
            if (parser().attribute_present("lowestEdited"))
            {
                target_.d_->file_version_.rup_build = string_to_size_t(parser().attribute("rupBuild"));
            }
            
            parser().next_expect(xml::parser::event_type::end_element);
        }
        else if (qname == xml::qname(xmlns_mc, "AlternateContent"))
        {
            parser().next_expect(xml::parser::event_type::start_element, xmlns_mc, "Choice");
            parser().content(xml::parser::content_type::complex);
            parser().attribute("Requires");
            parser().next_expect(xml::parser::event_type::start_element, xmlns_x15ac, "absPath");
            target_.set_absolute_path(path(parser().attribute("url")));
            parser().next_expect(xml::parser::event_type::end_element, xmlns_x15ac, "absPath");
            parser().next_expect(xml::parser::event_type::end_element, xmlns_mc, "Choice");
            parser().next_expect(xml::parser::event_type::end_element, xmlns_mc, "AlternateContent");
        }
        else if (qname == xml::qname(xmlns, "bookViews"))
        {
            if (parser().peek() == xml::parser::event_type::start_element)
            {
                parser().next_expect(xml::parser::event_type::start_element, xmlns, "workbookView");

                workbook_view view;
                view.x_window = string_to_size_t(parser().attribute("xWindow"));
                view.y_window = string_to_size_t(parser().attribute("yWindow"));
                view.window_width = string_to_size_t(parser().attribute("windowWidth"));
                view.window_height = string_to_size_t(parser().attribute("windowHeight"));
                
				if (parser().attribute_present("tabRatio"))
				{
					view.tab_ratio = string_to_size_t(parser().attribute("tabRatio"));
				}
                
                if (parser().attribute_present("activeTab"))
				{
					parser().attribute("activeTab");
				}
                
                if (parser().attribute_present("firstSheet"))
				{
					parser().attribute("firstSheet");
				}
                
                if (parser().attribute_present("showHorizontalScroll"))
				{
					parser().attribute("showHorizontalScroll");
				}
                
                if (parser().attribute_present("showSheetTabs"))
				{
					parser().attribute("showSheetTabs");
				}
                
                if (parser().attribute_present("showVerticalScroll"))
				{
					parser().attribute("showVerticalScroll");
				}
                
                target_.set_view(view);
                
                parser().next_expect(xml::parser::event_type::end_element, xmlns, "workbookView");
            }
            
            parser().next_expect(xml::parser::event_type::end_element, xmlns, "bookViews");
        }
        else if (qname == xml::qname(xmlns, "workbookPr"))
        {
            target_.d_->has_properties_ = true;

            if (parser().attribute_present("date1904"))
            {
                const auto value = parser().attribute("date1904");

                if (value == "1" || value == "true")
                {
                    target_.set_base_date(xlnt::calendar::mac_1904);
                }
            }

			if (parser().attribute_present("defaultThemeVersion"))
			{
				parser().attribute("defaultThemeVersion");
			}
            
            // todo: turn these structures into a method like skip_attribute(string name, bool optional)
            if (parser().attribute_present("backupFile"))
			{
				parser().attribute("backupFile");
			}
            
            if (parser().attribute_present("showObjects"))
			{
				parser().attribute("showObjects");
			}
            
            parser().next_expect(xml::parser::event_type::end_element, xmlns, "workbookPr");
        }
        else if (qname == xml::qname(xmlns, "sheets"))
        {
            std::size_t index = 0;
            
            while (true)
            {
                if (parser().peek() == xml::parser::event_type::end_element) break;
                
                parser().next_expect(xml::parser::event_type::start_element, xmlns_s, "sheet");

                std::string rel_id(parser().attribute(xml::qname(xmlns_r, "id")));
                std::string title(parser().attribute("name"));
                auto id = string_to_size_t(parser().attribute("sheetId"));

                sheet_title_id_map_[title] = id;
                sheet_title_index_map_[title] = index++;
                target_.d_->sheet_title_rel_id_map_[title] = rel_id;
                
                if (parser().attribute_present("state"))
				{
					parser().attribute("state");
				}
                
                parser().next_expect(xml::parser::event_type::end_element, xmlns_s, "sheet");
            }
            
            parser().next_expect(xml::parser::event_type::end_element, xmlns, "sheets");
        }
        else if (qname == xml::qname(xmlns, "calcPr"))
        {
            target_.d_->has_calculation_properties_ = true;

            if (parser().attribute_present("calcId"))
            {
                parser().attribute("calcId");
            }

			if (parser().attribute_present("concurrentCalc"))
			{
				parser().attribute("concurrentCalc");
			}
            
			if (parser().attribute_present("iterate"))
			{
				parser().attribute("iterate");
			}
            
			if (parser().attribute_present("iterateCount"))
			{
				parser().attribute("iterateCount");
			}
            
			if (parser().attribute_present("iterateDelta"))
			{
				parser().attribute("iterateDelta");
			}
            
			if (parser().attribute_present("refMode"))
			{
				parser().attribute("refMode");
			}

            parser().next_expect(xml::parser::event_type::end_element, xmlns, "calcPr");
        }
        else if (qname == xml::qname(xmlns, "extLst"))
        {
            parser().next_expect(xml::parser::event_type::start_element, xmlns, "ext");
            parser().content(xml::parser::content_type::complex);
            parser().attribute("uri");

			for (;;)
			{
				if (parser().peek() == xml::parser::event_type::end_element) break;

				parser().next_expect(xml::parser::event_type::start_element);

				if (parser().qname() == xml::qname(xmlns_mx, "ArchID"))
				{
					target_.d_->has_arch_id_ = true;
					parser().attribute("Flags");
				}
				else if (parser().qname() == xml::qname(xmlns_x15, "workbookPr"))
				{
					parser().attribute("chartTrackingRefBase");
				}
                
                if (parser().attribute_present("stringRefSyntax"))
                {
                    parser().attribute("stringRefSyntax");
                }

				parser().next_expect(xml::parser::event_type::end_element);
			}

            parser().next_expect(xml::parser::event_type::end_element, xmlns, "ext");
            parser().next_expect(xml::parser::event_type::end_element, xmlns, "extLst");
        }
        else if (qname == xml::qname(xmlns, "workbookProtection"))
        {
            while (parser().peek() != xml::parser::event_type::end_element
                || parser().qname() != xml::qname(xmlns, "workbookProtection"))
            {
                parser().next();
            }
            
            parser().next();
        }
    }
    
    parser().next_expect(xml::parser::event_type::end_element, xmlns, "workbook");
}

// Write Workbook Relationship Target Parts

void xlsx_consumer::read_calculation_chain()
{
}

void xlsx_consumer::read_chartsheet(const std::string &/*title*/)
{
}

void xlsx_consumer::read_connections()
{
}

void xlsx_consumer::read_custom_property()
{
}

void xlsx_consumer::read_custom_xml_mappings()
{
}

void xlsx_consumer::read_dialogsheet(const std::string &/*title*/)
{
}

void xlsx_consumer::read_external_workbook_references()
{
}

void xlsx_consumer::read_metadata()
{
}

void xlsx_consumer::read_pivot_table()
{
}

void xlsx_consumer::read_shared_string_table()
{
    static const auto &xmlns = constants::get_namespace("worksheet");
    
    parser().next_expect(xml::parser::event_type::start_element, xmlns, "sst");
	parser().content(xml::content::complex);

	std::size_t unique_count = 0;

	if (parser().attribute_present("uniqueCount"))
	{
		unique_count = string_to_size_t(parser().attribute("uniqueCount"));
	}

	auto &strings = target_.get_shared_strings();

    while (true)
    {
        if (parser().peek() == xml::parser::event_type::end_element) break;
        
        parser().next_expect(xml::parser::event_type::start_element, xmlns, "si");
		parser().content(xml::content::complex);
        parser().next_expect(xml::parser::event_type::start_element);
        
        formatted_text t;
        
        parser().attribute_map();
        
		if (parser().name() == "t")
		{
			parser().next_expect(xml::parser::event_type::characters);
			t.plain_text(parser().value());
		}
		else if (parser().name() == "r") // possible multiple text entities.
		{
            while (true)
            {
                if (parser().peek() == xml::parser::event_type::end_element) break;
                
                parser().next_expect(xml::parser::event_type::start_element, xmlns, "t");
				parser().next_expect(xml::parser::event_type::characters);
				text_run run;
                run.set_string(parser().value());

                if (parser().peek() == xml::parser::event_type::start_element)
                {
                    parser().next_expect(xml::parser::event_type::start_element, xmlns, "rPr");

                    while (true)
                    {
                        if (parser().peek() == xml::parser::event_type::end_element) break;
                        
                        parser().next_expect(xml::parser::event_type::start_element);

                        if (parser().qname() == xml::qname(xmlns, "sz"))
                        {
                            run.set_size(string_to_size_t(parser().attribute("val")));
                        }
                        else if (parser().qname() == xml::qname(xmlns, "rFont"))
                        {
                            run.set_font(parser().attribute("val"));
                        }
                        else if (parser().qname() == xml::qname(xmlns, "color"))
                        {
                            run.set_color(read_color(parser()));
                        }
                        else if (parser().qname() == xml::qname(xmlns, "family"))
                        {
                            run.set_family(string_to_size_t(parser().attribute("val")));
                        }
                        else if (parser().qname() == xml::qname(xmlns, "scheme"))
                        {
                            run.set_scheme(parser().attribute("val"));
                        }
                        
                        parser().next_expect(xml::parser::event_type::end_element);
                    }
                }

                t.add_run(run);
            }
		}

		parser().next_expect(xml::parser::event_type::end_element);
		parser().next_expect(xml::parser::event_type::end_element);

        strings.push_back(t);
	}

	if (unique_count != strings.size())
	{
		throw invalid_file("sizes don't match");
	}
}

void xlsx_consumer::read_shared_workbook_revision_headers()
{
}

void xlsx_consumer::read_shared_workbook()
{
}

void xlsx_consumer::read_shared_workbook_user_data()
{
}

void xlsx_consumer::read_stylesheet()
{
    static const auto &xmlns = constants::get_namespace("worksheet");
    static const auto &xmlns_mc = constants::get_namespace("mc");
    static const auto &xmlns_x14 = constants::get_namespace("x14");
    static const auto &xmlns_x14ac = constants::get_namespace("x14ac");
	static const auto &xmlns_x15 = constants::get_namespace("x15");
    
	auto &stylesheet = target_.impl().stylesheet_;

    parser().next_expect(xml::parser::event_type::start_element, xmlns, "styleSheet");
    parser().content(xml::parser::content_type::complex);

    while (true)
    {
        if (parser().peek() != xml::parser::event_type::start_namespace_decl) break;

        parser().next_expect(xml::parser::event_type::start_namespace_decl);

        if (parser().namespace_() == xmlns_x14ac)
        {
            target_.enable_x15();
        }
    }
    
    if (parser().attribute_present(xml::qname(xmlns_mc, "Ignorable")))
    {
        parser().attribute(xml::qname(xmlns_mc, "Ignorable"));
    }
    
    struct formatting_record
    {
        std::pair<class alignment, bool> alignment = { {}, 0 };
        std::pair<std::size_t, bool> border_id = { 0, false };
        std::pair<std::size_t, bool> fill_id = { 0, false };
        std::pair<std::size_t, bool> font_id = { 0, false };
        std::pair<std::size_t, bool> number_format_id = { 0, false };
        std::pair<class protection, bool> protection = { {}, false };
        std::pair<std::size_t, bool> style_id = { 0, false };
    };
    
    struct style_data
    {
        std::string name;
        std::size_t record_id;
        std::size_t builtin_id;
        bool custom_builtin;
    };

    std::vector<style_data> style_datas;
    std::vector<formatting_record> style_records;
    std::vector<formatting_record> format_records;

    while (true)
    {
        if (parser().peek() == xml::parser::event_type::end_element) break;

        parser().next_expect(xml::parser::event_type::start_element);
        parser().content(xml::parser::content_type::complex);

        if (parser().qname() == xml::qname(xmlns, "borders"))
        {
            stylesheet.borders.clear();
            
            auto count = parser().attribute<std::size_t>("count");
    
            while (true)
            {
                if (parser().peek() == xml::parser::event_type::end_element) break;

                stylesheet.borders.push_back(xlnt::border());
				auto &border = stylesheet.borders.back();

				parser().next_expect(xml::parser::event_type::start_element); // <border>
				parser().content(xml::parser::content_type::complex);
                
                auto diagonal = diagonal_direction::neither;
                
                if (parser().attribute_present("diagonalDown") && parser().attribute("diagonalDown") == "1")
                {
                    diagonal = diagonal_direction::down;
                }
                
                if (parser().attribute_present("diagonalUp") && parser().attribute("diagonalUp") == "1")
                {
                    diagonal = diagonal == diagonal_direction::down ? diagonal_direction::both : diagonal_direction::up;
                }
                
                if (diagonal != diagonal_direction::neither)
                {
                    border.diagonal(diagonal);
                }

				while (true)
				{
					if (parser().peek() == xml::parser::event_type::end_element) break;
					parser().next_expect(xml::parser::event_type::start_element);

					auto side_type = xml::value_traits<xlnt::border_side>::parse(parser().name(), parser());
					xlnt::border::border_property side;

					if (parser().attribute_present("style"))
					{
						side.style(parser().attribute<xlnt::border_style>("style"));
					}

					if (parser().peek() == xml::parser::event_type::start_element)
					{
						parser().next_expect(xml::parser::event_type::start_element, "color");
						side.color(read_color(parser()));
						parser().next_expect(xml::parser::event_type::end_element, "color");
					}

					border.side(side_type, side);
					parser().next_expect(xml::parser::event_type::end_element);
				}

				parser().next_expect(xml::parser::event_type::end_element); // </border>
            }
                        
            if (count != stylesheet.borders.size())
            {
                throw xlnt::exception("counts don't match");
            }
        }
        else if (parser().qname() == xml::qname(xmlns, "fills"))
        {
            stylesheet.fills.clear();
            
            auto count = parser().attribute<std::size_t>("count");

            while (true)
            {
                if (parser().peek() == xml::parser::event_type::end_element) break;

                stylesheet.fills.push_back(xlnt::fill());
				auto &new_fill = stylesheet.fills.back();

				parser().next_expect(xml::parser::event_type::start_element, xmlns, "fill");
				parser().content(xml::parser::content_type::complex);
				parser().next_expect(xml::parser::event_type::start_element);

				if (parser().qname() == xml::qname(xmlns, "patternFill"))
				{
					xlnt::pattern_fill pattern;

					if (parser().attribute_present("patternType"))
					{
						pattern.type(parser().attribute<xlnt::pattern_fill_type>("patternType"));

						while (true)
						{
							if (parser().peek() == xml::parser::event_type::end_element)
							{
								break;
							}

							parser().next_expect(xml::parser::event_type::start_element);

							if (parser().name() == "fgColor")
							{
								pattern.foreground(read_color(parser()));
							}
							else if (parser().name() == "bgColor")
							{
								pattern.background(read_color(parser()));
							}

							parser().next_expect(xml::parser::event_type::end_element);
						}
					}

					new_fill = pattern;
				}
				else if (parser().qname() == xml::qname(xmlns, "gradientFill"))
				{
					xlnt::gradient_fill gradient;

					if (parser().attribute_present("type"))
					{
						gradient.type(parser().attribute<xlnt::gradient_fill_type>("type"));
					}
					else
					{
						gradient.type(xlnt::gradient_fill_type::linear);
					}

					while (true)
					{
						if (parser().peek() == xml::parser::event_type::end_element) break;

						parser().next_expect(xml::parser::event_type::start_element, "stop");
						auto position = parser().attribute<double>("position");
						parser().next_expect(xml::parser::event_type::start_element, "color");
						auto color = read_color(parser());
						parser().next_expect(xml::parser::event_type::end_element, "color");
						parser().next_expect(xml::parser::event_type::end_element, "stop");

						gradient.add_stop(position, color);
					}

					new_fill = gradient;
				}

				parser().next_expect(xml::parser::event_type::end_element); // </gradientFill> or </patternFill>
				parser().next_expect(xml::parser::event_type::end_element); // </fill>
            }
                        
            if (count != stylesheet.fills.size())
            {
                throw xlnt::exception("counts don't match");
            }
        }
        else if (parser().qname() == xml::qname(xmlns, "fonts"))
        {
            stylesheet.fonts.clear();

            auto count = parser().attribute<std::size_t>("count");
            
            if (parser().attribute_present(xml::qname(xmlns_x14ac, "knownFonts")))
            {
                parser().attribute(xml::qname(xmlns_x14ac, "knownFonts"));
            }

            while (true)
            {
                if (parser().peek() == xml::parser::event_type::end_element) break;

				stylesheet.fonts.push_back(xlnt::font());
				auto &new_font = stylesheet.fonts.back();

				parser().next_expect(xml::parser::event_type::start_element, xmlns, "font");
				parser().content(xml::parser::content_type::complex);

				while (true)
				{
					if (parser().peek() == xml::parser::event_type::end_element) break;

					parser().next_expect(xml::parser::event_type::start_element);
					parser().content(xml::parser::content_type::simple);

					if (parser().name() == "sz")
					{
						new_font.size(string_to_size_t(parser().attribute("val")));
					}
					else if (parser().name() == "name")
					{
						new_font.name(parser().attribute("val"));
					}
					else if (parser().name() == "color")
					{
						new_font.color(read_color(parser()));
					}
					else if (parser().name() == "family")
					{
						new_font.family(string_to_size_t(parser().attribute("val")));
					}
					else if (parser().name() == "scheme")
					{
						new_font.scheme(parser().attribute("val"));
					}
					else if (parser().name() == "b")
					{
						if (parser().attribute_present("val"))
						{
							new_font.bold(is_true(parser().attribute("val")));
						}
						else
						{
							new_font.bold(true);
						}
					}
					else if (parser().name() == "vertAlign")
					{
						new_font.superscript(parser().attribute("val") == "superscript");
					}
					else if (parser().name() == "strike")
					{
						if (parser().attribute_present("val"))
						{
							new_font.strikethrough(is_true(parser().attribute("val")));
						}
						else
						{
							new_font.strikethrough(true);
						}
					}
					else if (parser().name() == "i")
					{
						if (parser().attribute_present("val"))
						{
							new_font.italic(is_true(parser().attribute("val")));
						}
						else
						{
							new_font.italic(true);
						}
					}
					else if (parser().name() == "u")
					{
						if (parser().attribute_present("val"))
						{
							new_font.underline(parser().attribute<xlnt::font::underline_style>("val"));
						}
						else
						{
							new_font.underline(xlnt::font::underline_style::single);
						}
					}
                    else if (parser().name() == "charset")
					{
						if (parser().attribute_present("val"))
						{
							parser().attribute("val");
						}
					}

					parser().next_expect(xml::parser::event_type::end_element);
				}

				parser().next_expect(xml::parser::event_type::end_element, xmlns, "font");
            }
            
            if (count != stylesheet.fonts.size())
            {
                throw xlnt::exception("counts don't match");
            }
        }
        else if (parser().qname() == xml::qname(xmlns, "numFmts"))
        {
            stylesheet.number_formats.clear();
            
            auto count = parser().attribute<std::size_t>("count");

            while (true)
            {
                if (parser().peek() == xml::parser::event_type::end_element) break;
                parser().next_expect(xml::parser::event_type::start_element, xmlns, "numFmt");

                auto format_string = parser().attribute("formatCode");

                if (format_string == "GENERAL")
                {
                    format_string = "General";
                }

                xlnt::number_format nf;

                nf.set_format_string(format_string);
                nf.set_id(string_to_size_t(parser().attribute("numFmtId")));

                stylesheet.number_formats.push_back(nf);
                parser().next_expect(xml::parser::event_type::end_element); // numFmt
            }
            
            if (count != stylesheet.number_formats.size())
            {
                throw xlnt::exception("counts don't match");
            }
        }
        else if (parser().qname() == xml::qname(xmlns, "colors"))
        {
            while (parser().peek() != xml::parser::event_type::end_element
                || parser().qname() != xml::qname(xmlns, "colors"))
            {
                if (parser().next() == xml::parser::event_type::start_element)
                {
                    parser().attribute_map();
                }
            }
        }
        else if (parser().qname() == xml::qname(xmlns, "cellStyles"))
        {
            auto count = parser().attribute<std::size_t>("count");
            
            while (true)
            {
                if (parser().peek() == xml::parser::event_type::end_element) break;
                
                auto &data = *style_datas.emplace(style_datas.end());
                
                parser().next_expect(xml::parser::event_type::start_element, xmlns, "cellStyle");

                data.name = parser().attribute("name");
                data.record_id = parser().attribute<std::size_t>("xfId");
                data.builtin_id = parser().attribute<std::size_t>("builtinId");

                if (parser().attribute_present("customBuiltin"))
                {
                    data.custom_builtin = is_true(parser().attribute("customBuiltin"));
                }

                parser().next_expect(xml::parser::event_type::end_element, xmlns, "cellStyle");
            }
            
            if (count != style_datas.size())
            {
                throw xlnt::exception("counts don't match");
            }
        }
        else if (parser().qname() == xml::qname(xmlns, "cellStyleXfs")
            || parser().qname() == xml::qname(xmlns, "cellXfs"))
        {
            auto in_style_records = parser().name() == "cellStyleXfs";
            auto count = parser().attribute<std::size_t>("count");
            
            while (true)
            {
                if (parser().peek() == xml::parser::event_type::end_element) break;
                
                parser().next_expect(xml::parser::event_type::start_element, xmlns, "xf");
				parser().content(xml::content::complex);

                auto &record = *(!in_style_records
                    ? format_records.emplace(format_records.end())
                    : style_records.emplace(style_records.end()));

                auto apply_alignment_present = parser().attribute_present("applyAlignment");
                auto alignment_applied = apply_alignment_present
                    && is_true(parser().attribute("applyAlignment"));
                record.alignment.second = alignment_applied;

                auto border_applied = parser().attribute_present("applyBorder")
                    && is_true(parser().attribute("applyBorder"));
                auto border_index = parser().attribute_present("borderId")
                    ? string_to_size_t(parser().attribute("borderId")) : 0;
                record.border_id = { border_index, border_applied };
                
                auto fill_applied = parser().attribute_present("applyFill")
                    && is_true(parser().attribute("applyFill"));
                auto fill_index = parser().attribute_present("fillId")
                    ? string_to_size_t(parser().attribute("fillId")) : 0;
                record.fill_id = { fill_index, fill_applied };

                auto font_applied = parser().attribute_present("applyFont")
                    && is_true(parser().attribute("applyFont"));
                auto font_index = parser().attribute_present("fontId")
                    ? string_to_size_t(parser().attribute("fontId")) : 0;
                record.font_id = { font_index, font_applied };

                auto number_format_applied = parser().attribute_present("applyNumberFormat")
                    && is_true(parser().attribute("applyNumberFormat"));
                auto number_format_id = parser().attribute_present("numFmtId")
                    ? string_to_size_t(parser().attribute("numFmtId")) : 0;
                record.number_format_id = { number_format_id, number_format_applied };

                auto apply_protection_present = parser().attribute_present("applyProtection");
                auto protection_applied = apply_protection_present
                    && is_true(parser().attribute("applyProtection"));
                record.protection.second = protection_applied;

                if (parser().attribute_present("xfId") && parser().name() == "cellXfs")
                {
                    record.style_id = { parser().attribute<std::size_t>("xfId"), true };
                }
                
                while (true)
                {
                    if (parser().peek() == xml::parser::event_type::end_element) break;
                    
                    parser().next_expect(xml::parser::event_type::start_element);
                    
                    if (parser().qname() == xml::qname(xmlns, "alignment"))
                    {
						if (parser().attribute_present("wrapText"))
						{
							record.alignment.first.wrap(is_true(parser().attribute("wrapText")));
						}

						if (parser().attribute_present("shrinkToFit"))
						{
							record.alignment.first.shrink(is_true(parser().attribute("shrinkToFit")));
						}

						if (parser().attribute_present("indent"))
						{
							record.alignment.first.indent(parser().attribute<int>("indent"));
						}

						if (parser().attribute_present("textRotation"))
						{
							record.alignment.first.rotation(parser().attribute<int>("textRotation"));
						}

						if (parser().attribute_present("vertical"))
						{
							record.alignment.first.vertical(
								parser().attribute<xlnt::vertical_alignment>("vertical"));
						}

						if (parser().attribute_present("horizontal"))
						{
							record.alignment.first.horizontal(
								parser().attribute<xlnt::horizontal_alignment>("horizontal"));
						}

                        record.alignment.second = !apply_alignment_present || alignment_applied;
                    }
                    else if (parser().qname() == xml::qname(xmlns, "protection"))
                    {
						record.protection.first.locked(is_true(parser().attribute("locked")));
						record.protection.first.hidden(is_true(parser().attribute("hidden")));
                        record.protection.second = !apply_protection_present || protection_applied;
                    }
                    
                    parser().next_expect(xml::parser::event_type::end_element, parser().qname());
                }
                
                parser().next_expect(xml::parser::event_type::end_element, xmlns, "xf");
                
            }
            
            if ((in_style_records && count != style_records.size())
                || (!in_style_records && count != format_records.size()))
            {
                throw xlnt::exception("counts don't match");
            }
        }
        else if (parser().qname() == xml::qname(xmlns, "dxfs"))
        {
            auto count = parser().attribute<std::size_t>("count");
            std::size_t processed = 0;
            
            while (true)
            {
                if (parser().peek() == xml::parser::event_type::end_element) break;
                parser().next_expect(xml::parser::event_type::start_element);
                parser().next_expect(xml::parser::event_type::end_element);
            }
            
            if (count != processed)
            {
                throw xlnt::exception("counts don't match");
            }
        }
        else if (parser().qname() == xml::qname(xmlns, "tableStyles"))
        {
            auto default_table_style = parser().attribute("defaultTableStyle");
            auto default_pivot_style = parser().attribute("defaultPivotStyle");
            auto count = parser().attribute<std::size_t>("count");
            std::size_t processed = 0;
            
            while (true)
            {
                if (parser().peek() == xml::parser::event_type::end_element) break;
                parser().next_expect(xml::parser::event_type::start_element);
                parser().next_expect(xml::parser::event_type::end_element);
            }
            
            if (count != processed)
            {
                throw xlnt::exception("counts don't match");
            }
        }
        else if (parser().qname() == xml::qname(xmlns, "extLst"))
        {
			for (;;)
			{
				if (parser().peek() == xml::parser::event_type::end_element) break;

				parser().next_expect(xml::parser::event_type::start_element, xmlns, "ext");
				parser().content(xml::parser::content_type::complex);
				parser().attribute("uri");
				parser().next_expect(xml::parser::event_type::start_namespace_decl);

				parser().next_expect(xml::parser::event_type::start_element);

				if (parser().qname() == xml::qname(xmlns_x14, "slicerStyles"))
				{
					parser().attribute("defaultSlicerStyle");
				}
				else if (parser().qname() == xml::qname(xmlns_x15, "timelineStyles"))
				{
					parser().attribute("defaultTimelineStyle");
				}

				parser().next_expect(xml::parser::event_type::end_element);

				parser().next_expect(xml::parser::event_type::end_element, xmlns, "ext");
				parser().next_expect(xml::parser::event_type::end_namespace_decl);
			}
        }
        else if (parser().qname() == xml::qname(xmlns, "indexedColors"))
        {
            parser().next_expect(xml::parser::event_type::end_element, xmlns, "indexedColors");
        }
        
        parser().next_expect(xml::parser::event_type::end_element);
    }
    
    parser().next_expect(xml::parser::event_type::end_element);
    
    auto lookup_number_format = [&](std::size_t number_format_id)
    {
        auto result = number_format::general();
        bool is_custom_number_format = false;
        
        for (const auto &nf : stylesheet.number_formats)
        {
            if (nf.get_id() == number_format_id)
            {
                result = nf;
                is_custom_number_format = true;
                break;
            }
        }

        if (number_format_id < 164 && !is_custom_number_format)
        {
            result = number_format::from_builtin_id(number_format_id);
        }
        
        return result;
    };
    
    std::size_t xf_id = 0;
    
    for (const auto &record : style_records)
    {
        auto style_data_iter = std::find_if(style_datas.begin(), style_datas.end(),
            [&xf_id](const style_data &s) { return s.record_id == xf_id; });
        ++xf_id;
        
        if (style_data_iter == style_datas.end()) continue;

        auto new_style = stylesheet.create_style(style_data_iter->name);
        new_style.builtin_id(style_data_iter->builtin_id);

        new_style.alignment(record.alignment.first, record.alignment.second);
        new_style.border(stylesheet.borders.at(record.border_id.first), record.border_id.second);
        new_style.fill(stylesheet.fills.at(record.fill_id.first), record.fill_id.second);
        new_style.font(stylesheet.fonts.at(record.font_id.first), record.font_id.second);
        new_style.number_format(lookup_number_format(record.number_format_id.first), record.number_format_id.second);
        new_style.protection(record.protection.first, record.protection.second);
    }
    
    std::size_t record_index = 0;

    for (const auto &record : format_records)
    {
        stylesheet.format_impls.push_back(format_impl());
        auto &new_format = stylesheet.format_impls.back();
        
        new_format.id = record_index++;
        new_format.parent = &stylesheet;

        ++new_format.references;

        if (record.style_id.second)
        {
            new_format.style = stylesheet.style_names[record.style_id.first];
        }

        new_format.alignment_id = stylesheet.find_or_add(stylesheet.alignments, record.alignment.first);
        new_format.alignment_applied = record.alignment.second;
        new_format.border_id = record.border_id.first;
        new_format.border_applied = record.border_id.second;
        new_format.fill_id = record.fill_id.first;
        new_format.fill_applied = record.fill_id.second;
        new_format.font_id = record.font_id.first;
        new_format.font_applied = record.font_id.second;
        new_format.number_format_id = record.number_format_id.first;
        new_format.number_format_applied = record.number_format_id.second;
        new_format.protection_id = stylesheet.find_or_add(stylesheet.protections, record.protection.first);
        new_format.protection_applied = record.protection.second;
    }
}

void xlsx_consumer::read_theme()
{
	target_.set_theme(theme());
}

void xlsx_consumer::read_volatile_dependencies()
{
}

void xlsx_consumer::read_worksheet(const std::string &rel_id)
{
    static const auto &xmlns = constants::get_namespace("worksheet");
    static const auto &xmlns_mc = constants::get_namespace("mc");
    static const auto &xmlns_r = constants::get_namespace("r");
    static const auto &xmlns_x14ac = constants::get_namespace("x14ac");

	auto title = std::find_if(target_.d_->sheet_title_rel_id_map_.begin(),
		target_.d_->sheet_title_rel_id_map_.end(),
		[&](const std::pair<std::string, std::string> &p)
	{
		return p.second == rel_id;
	})->first;

	auto id = sheet_title_id_map_[title];
	auto index = sheet_title_index_map_[title];

	auto insertion_iter = target_.d_->worksheets_.begin();
	while (insertion_iter != target_.d_->worksheets_.end()
		&& sheet_title_index_map_[insertion_iter->title_] < index)
	{
		++insertion_iter;
	}

	target_.d_->worksheets_.emplace(insertion_iter, &target_, id, title);

	auto ws = target_.get_sheet_by_id(id);

	parser().next_expect(xml::parser::event_type::start_element, xmlns, "worksheet");
    parser().content(xml::parser::content_type::complex);

    while (parser().peek() == xml::parser::event_type::start_namespace_decl)
	{
        parser().next_expect(xml::parser::event_type::start_namespace_decl);

        if (parser().namespace_() == xmlns_x14ac)
        {
            ws.enable_x14ac();
        }
	}
    
    if (parser().attribute_present(xml::qname(xmlns_mc, "Ignorable")))
    {
        parser().attribute(xml::qname(xmlns_mc, "Ignorable"));
    }

	xlnt::range_reference full_range;

    while (true)
    {
        if (parser().peek() == xml::parser::event_type::end_element) break;
        
        parser().next_expect(xml::parser::event_type::start_element);
        parser().content(xml::parser::content_type::complex);
        
        if (parser().qname() == xml::qname(xmlns, "dimension"))
        {
            full_range = xlnt::range_reference(parser().attribute("ref"));
            ws.d_->has_dimension_ = true;
            parser().next_expect(xml::parser::event_type::end_element, xmlns, "dimension");
        }
        else if (parser().qname() == xml::qname(xmlns, "sheetViews"))
        {
            ws.d_->has_view_ = true;
            
            while (true)
            {
                parser().attribute_map();

                if (parser().next() == xml::parser::event_type::end_element && parser().name() == "sheetViews")
                {
                    break;
                }
            }
        }
        else if (parser().qname() == xml::qname(xmlns, "sheetFormatPr"))
        {
            ws.d_->has_format_properties_ = true;
            
            while (true)
            {
                parser().attribute_map();

                if (parser().next() == xml::parser::event_type::end_element && parser().name() == "sheetFormatPr")
                {
                    break;
                }
            }
        }
        else if (parser().qname() == xml::qname(xmlns, "mergeCells"))
        {
            auto count = std::stoull(parser().attribute("count"));

            while (true)
            {
                if (parser().peek() == xml::parser::event_type::end_element) break;

                parser().next_expect(xml::parser::event_type::start_element, xmlns, "mergeCell");
                ws.merge_cells(range_reference(parser().attribute("ref")));
                parser().next_expect(xml::parser::event_type::start_element, xmlns, "mergeCell");

                count--;
            }

            if (count != 0)
            {
                throw invalid_file("sizes don't match");
            }
            
            parser().next_expect(xml::parser::event_type::end_element, xmlns, "mergeCells");
        }
        else if (parser().qname() == xml::qname(xmlns, "sheetData"))
        {
            auto &shared_strings = target_.get_shared_strings();

            while (true)
            {
                if (parser().peek() == xml::parser::event_type::end_element) break;
                
                parser().next_expect(xml::parser::event_type::start_element, xmlns, "row");
				parser().content(xml::content::complex);

                auto row_index = static_cast<row_t>(std::stoull(parser().attribute("r")));

                if (parser().attribute_present("ht"))
                {
                    ws.get_row_properties(row_index).height = std::stold(parser().attribute("ht"));
                }

				if (parser().attribute_present("customHeight"))
				{
                    auto custom_height = parser().attribute("customHeight");
                    
                    if (custom_height != "false")
                    {
                        ws.get_row_properties(row_index).height = std::stold(custom_height);
                    }
				}

				if (parser().attribute_present(xml::qname(xmlns_x14ac, "dyDescent")))
				{
					parser().attribute(xml::qname(xmlns_x14ac, "dyDescent"));
				}

                auto min_column = full_range.get_top_left().get_column_index();
                auto max_column = full_range.get_bottom_right().get_column_index();

                if (parser().attribute_present("spans"))
                {
                    std::string span_string = parser().attribute("spans");
                    auto colon_index = span_string.find(':');

                    if (colon_index != std::string::npos)
                    {
                        min_column = static_cast<column_t::index_t>(std::stoll(span_string.substr(0, colon_index)));
                        max_column = static_cast<column_t::index_t>(std::stoll(span_string.substr(colon_index + 1)));
                    }
                }

				if (parser().attribute_present("customFormat"))
				{
					parser().attribute("customFormat");
				}

				if (parser().attribute_present("s"))
				{
					parser().attribute("s");
				}
                
                if (parser().attribute_present("collapsed"))
				{
					parser().attribute("collapsed");
				}
                
                if (parser().attribute_present("hidden"))
				{
					parser().attribute("hidden");
				}
                
                if (parser().attribute_present("outlineLevel"))
				{
					parser().attribute("outlineLevel");
				}

                while (true)
                {
                    if (parser().peek() == xml::parser::event_type::end_element) break;

                    parser().next_expect(xml::parser::event_type::start_element, xmlns, "c");
					parser().content(xml::content::complex);
                    auto cell = ws.get_cell(cell_reference(parser().attribute("r")));
                    
                    auto has_type = parser().attribute_present("t");
                    auto type = has_type ? parser().attribute("t") : "";

                    auto has_format = parser().attribute_present("s");
                    auto format_id = static_cast<std::size_t>(has_format ? std::stoull(parser().attribute("s")) : 0LL);
    
                    auto has_value = false;
                    auto value_string = std::string();
                    
                    auto has_formula = false;
                    auto has_shared_formula = false;
                    auto formula_value_string = std::string();
    
                    while (true)
                    {
                        if (parser().peek() == xml::parser::event_type::end_element) break;
                        
                        parser().next_expect(xml::parser::event_type::start_element);
                        
                        if (parser().qname() == xml::qname(xmlns, "v"))
                        {
                            has_value = true;

                            // <v> might be empty, check first
                            if (parser().peek() == xml::parser::event_type::characters)
                            {
                                parser().next_expect(xml::parser::event_type::characters);
                                value_string = parser().value();
                            }
                        }
                        else if (parser().qname() == xml::qname(xmlns, "f"))
                        {
                            has_formula = true;
                            has_shared_formula = parser().attribute_present("t")
								&& parser().attribute("t") == "shared";
							parser().next_expect(xml::parser::event_type::characters);
                            formula_value_string = parser().value();
                        }
                        else if (parser().qname() == xml::qname(xmlns, "is"))
                        {
                            parser().next_expect(xml::parser::event_type::start_element, xmlns, "t");
							parser().next_expect(xml::parser::event_type::characters);
                            value_string = parser().value();
                            parser().next_expect(xml::parser::event_type::end_element, xmlns, "t");
                        }
                        
                        parser().next_expect(xml::parser::event_type::end_element);
                    }

                    if (has_formula && !has_shared_formula && !ws.get_workbook().get_data_only())
                    {
                        cell.set_formula(formula_value_string);
                    }

                    if (has_type && (type == "inlineStr" || type =="str"))
                    {
                        cell.set_value(value_string);
                    }
                    else if (has_type && type == "s" && !has_formula)
                    {
                        auto shared_string_index = static_cast<std::size_t>(std::stoull(value_string));
                        auto shared_string = shared_strings.at(shared_string_index);
                        cell.set_value(shared_string);
                    }
                    else if (has_type && type == "b") // boolean
                    {
                        cell.set_value(value_string != "0");
                    }
                    else if (has_value && !value_string.empty())
                    {
                        if (!value_string.empty() && value_string[0] == '#')
                        {
                            cell.set_error(value_string);
                        }
                        else
                        {
                            cell.set_value(std::stold(value_string));
                        }
                    }

                    if (has_format)
                    {
                        cell.format(target_.get_format(format_id));
                    }
                    
                    parser().next_expect(xml::parser::event_type::end_element, xmlns, "c");
                }

				parser().next_expect(xml::parser::event_type::end_element, xmlns, "row");
            }
            
            parser().next_expect(xml::parser::event_type::end_element, xmlns, "sheetData");
        }
        else if (parser().qname() == xml::qname(xmlns, "cols"))
        {
            while (true)
            {
                if (parser().peek() == xml::parser::event_type::end_element) break;
                
                parser().next_expect(xml::parser::event_type::start_element, xmlns, "col");

                auto min = static_cast<column_t::index_t>(std::stoull(parser().attribute("min")));
                auto max = static_cast<column_t::index_t>(std::stoull(parser().attribute("max")));

                auto width = std::stold(parser().attribute("width"));
                auto column_style = parser().attribute_present("style") 
                    ? parser().attribute<std::size_t>("style") : static_cast<std::size_t>(0);
                auto custom = parser().attribute_present("customWidth")
                    ? is_true(parser().attribute("customWidth")) : false;

                for (auto column = min; column <= max; column++)
                {
                    column_properties props;

                    props.width = width;
                    props.style = column_style;
                    props.custom = custom;

                    ws.add_column_properties(column, props);
                }

				if (parser().attribute_present("bestFit"))
				{
					parser().attribute("bestFit");
				}
                
				if (parser().attribute_present("collapsed"))
				{
					parser().attribute("collapsed");
				}
                
                if (parser().attribute_present("hidden"))
				{
					parser().attribute("hidden");
				}
                
                if (parser().attribute_present("outlineLevel"))
				{
					parser().attribute("outlineLevel");
				}
                
                parser().next_expect(xml::parser::event_type::end_element, xmlns, "col");
            }
            
            parser().next_expect(xml::parser::event_type::end_element, xmlns, "cols");
        }
        else if (parser().qname() == xml::qname(xmlns, "autoFilter"))
        {
            ws.auto_filter(xlnt::range_reference(parser().attribute("ref")));
            parser().next_expect(xml::parser::event_type::end_element, xmlns, "autoFilter");
        }
        else if (parser().qname() == xml::qname(xmlns, "pageMargins"))
        {
            page_margins margins;

            margins.set_top(parser().attribute<double>("top"));
            margins.set_bottom(parser().attribute<double>("bottom"));
            margins.set_left(parser().attribute<double>("left"));
            margins.set_right(parser().attribute<double>("right"));
            margins.set_header(parser().attribute<double>("header"));
            margins.set_footer(parser().attribute<double>("footer"));

            ws.set_page_margins(margins);
            
            parser().next_expect(xml::parser::event_type::end_element, xmlns, "pageMargins");
        }
        else if (parser().qname() == xml::qname(xmlns, "pageSetUpPr"))
        {
            parser().attribute_map();
            parser().next_expect(xml::parser::event_type::end_element, xmlns, "pageSetUpPr");
        }
        else if (parser().qname() == xml::qname(xmlns, "printOptions"))
        {
            parser().attribute_map();
            parser().next_expect(xml::parser::event_type::end_element, xmlns, "printOptions");
        }
        else if (parser().qname() == xml::qname(xmlns, "pageSetup"))
        {
            parser().attribute_map();
            parser().next_expect(xml::parser::event_type::end_element, xmlns, "pageSetup");
        }
        else if (parser().qname() == xml::qname(xmlns, "sheetPr"))
        {
            parser().attribute_map();

            while (parser().peek() != xml::parser::event_type::end_element
                || parser().qname() != xml::qname(xmlns, "sheetPr"))
            {
                if (parser().next() == xml::parser::event_type::start_element)
                {
                    parser().attribute_map();
                }
            }
            
            parser().next();
        }
        else if (parser().qname() == xml::qname(xmlns, "headerFooter"))
        {
            parser().attribute_map();

            while (parser().peek() != xml::parser::event_type::end_element
                || parser().qname() != xml::qname(xmlns, "headerFooter"))
            {
                if (parser().next() == xml::parser::event_type::start_element)
                {
                    parser().attribute_map();
                }
            }
            
            parser().next();
        }
        else if (parser().qname() == xml::qname(xmlns, "legacyDrawing"))
        {
	    parser().attribute(xml::qname(xmlns_r, "id"));
            parser().next_expect(xml::parser::event_type::end_element, xmlns, "legacyDrawing");
	}
    }
    
    parser().next_expect(xml::parser::event_type::end_element, xmlns, "worksheet");

    auto &manifest = target_.get_manifest();
    const auto workbook_rel = manifest.get_relationship(path("/"), relationship::type::office_document);
    const auto sheet_rel = manifest.get_relationship(workbook_rel.get_target().get_path(), rel_id);
    path sheet_path(sheet_rel.get_source().get_path().parent().append(sheet_rel.get_target().get_path()));

    if (manifest.has_relationship(sheet_path, xlnt::relationship::type::comments))
    {
        auto comments_part = manifest.canonicalize({workbook_rel, sheet_rel,
            manifest.get_relationship(sheet_path, xlnt::relationship::type::comments)});

        auto receive = xml::parser::receive_default;
        xml::parser parser(archive_->open(comments_part.string()), comments_part.string(), receive);
        parser_ = &parser;
        
        read_comments(ws);
        
        if (manifest.has_relationship(sheet_path, xlnt::relationship::type::vml_drawing))
        {
            auto vml_drawings_part = manifest.canonicalize({workbook_rel, sheet_rel,
                manifest.get_relationship(sheet_path, xlnt::relationship::type::vml_drawing)});

            auto receive = xml::parser::receive_default;
            xml::parser parser(archive_->open(vml_drawings_part.string()), vml_drawings_part.string(), receive);
            parser_ = &parser;
            
            read_vml_drawings(ws);
        }
    }
}

// Sheet Relationship Target Parts

void xlsx_consumer::read_vml_drawings(worksheet ws)
{

}

void xlsx_consumer::read_comments(worksheet ws)
{
    static const auto &xmlns = xlnt::constants::get_namespace("worksheet");
    
    std::vector<std::string> authors;
    
    parser().next_expect(xml::parser::event_type::start_element, xmlns, "comments");
	parser().content(xml::content::complex);
    parser().next_expect(xml::parser::event_type::start_element, xmlns, "authors");
	parser().content(xml::content::complex);
    
    for (;;)
    {
        if (parser().peek() == xml::parser::event_type::end_element) break;

        parser().next_expect(xml::parser::event_type::start_element, xmlns, "author");
        parser().next_expect(xml::parser::event_type::characters);
        authors.push_back(parser().value());
        parser().next_expect(xml::parser::event_type::end_element, xmlns, "author");
    }
    
    parser().next_expect(xml::parser::event_type::end_element, xmlns, "authors");

    parser().next_expect(xml::parser::event_type::start_element, xmlns, "commentList");
	parser().content(xml::content::complex);
    
    for (;;)
    {
        if (parser().peek() == xml::parser::event_type::end_element) break;

        parser().next_expect(xml::parser::event_type::start_element, xmlns, "comment");
		parser().content(xml::content::complex);

        auto cell_ref = parser().attribute("ref");
        auto author_id = parser().attribute<std::size_t>("authorId");
        
        parser().next_expect(xml::parser::event_type::start_element, xmlns, "text");
		parser().content(xml::content::complex);

        // todo: this is duplicated from shared strings
        formatted_text text;

        for (;;)
        {
            if (parser().peek() == xml::parser::event_type::end_element) break;
            
            parser().next_expect(xml::parser::event_type::start_element, xmlns, "r");
			parser().content(xml::content::complex);

            text_run run;
            
            for (;;)
            {
                if (parser().peek() == xml::parser::event_type::end_element) break;
                
                parser().next_expect(xml::parser::event_type::start_element);
                
                if (parser().name() == "t")
                {
                    parser().next_expect(xml::parser::event_type::characters);
                    run.set_string(parser().value());
                    parser().next_expect(xml::parser::event_type::end_element, xmlns, "t");
                }
                else if (parser().name() == "rPr")
                {
					parser().content(xml::content::complex);

                    for (;;)
                    {
                        if (parser().peek() == xml::parser::event_type::end_element) break;
                        
                        parser().next_expect(xml::parser::event_type::start_element);

                        if (parser().qname() == xml::qname(xmlns, "sz"))
                        {
                            run.set_size(string_to_size_t(parser().attribute("val")));
                        }
                        else if (parser().qname() == xml::qname(xmlns, "rFont"))
                        {
                            run.set_font(parser().attribute("val"));
                        }
                        else if (parser().qname() == xml::qname(xmlns, "color"))
                        {
                            run.set_color(read_color(parser()));
                        }
                        else if (parser().qname() == xml::qname(xmlns, "family"))
                        {
                            run.set_family(string_to_size_t(parser().attribute("val")));
                        }
                        else if (parser().qname() == xml::qname(xmlns, "scheme"))
                        {
                            run.set_scheme(parser().attribute("val"));
                        }
                        else if (parser().qname() == xml::qname(xmlns, "b"))
                        {
                            run.set_bold(true);
                        }
                        
                        parser().next_expect(xml::parser::event_type::end_element);
                    }
                    
                    parser().next_expect(xml::parser::event_type::end_element, xmlns, "rPr");
                }
            }
            
            text.add_run(run);
            parser().next_expect(xml::parser::event_type::end_element, xmlns, "r");
        }
    
        ws.get_cell(cell_ref).comment(comment(text, authors.at(author_id)));
        parser().next_expect(xml::parser::event_type::end_element, xmlns, "text");
        parser().next_expect(xml::parser::event_type::end_element, xmlns, "comment");
    }
    
    parser().next_expect(xml::parser::event_type::end_element, xmlns, "commentList");
    parser().next_expect(xml::parser::event_type::end_element, xmlns, "comments");
}

void xlsx_consumer::read_drawings()
{
}

// Unknown Parts

void xlsx_consumer::read_unknown_parts()
{
}

void xlsx_consumer::read_unknown_relationships()
{
}

} // namespace detail
} // namepsace xlnt
