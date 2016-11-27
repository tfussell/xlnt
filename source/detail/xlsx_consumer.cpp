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

#include <detail/constants.hpp>
#include <detail/custom_value_traits.hpp>
#include <detail/workbook_impl.hpp>
#include <detail/xlsx_consumer.hpp>
#include <detail/zip.hpp>
#include <xlnt/cell/cell.hpp>
#include <xlnt/packaging/manifest.hpp>
#include <xlnt/utils/path.hpp>
#include <xlnt/workbook/const_worksheet_iterator.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/worksheet/worksheet.hpp>

namespace std {

/// <summary>
/// Allows xml::qname to be used as a key in a std::unordered_map.
/// </summary>
template <>
struct hash<xml::qname>
{
    std::size_t operator()(const xml::qname &k) const
    {
        static std::hash<std::string> hasher;
        return hasher(k.string());
    }
};

} // namespace std

namespace {

#define THROW_ON_INVALID_XML

/// <summary>
/// Returns true if bool_string represents a true xsd:boolean.
/// </summary>
bool is_true(const std::string &bool_string)
{
    if (bool_string == "1" || bool_string == "true")
    {
        return true;
    }

#ifdef THROW_ON_INVALID_XML
    if (bool_string == "0" || bool_string == "false")
    {
        return false;
    }

    throw xlnt::exception("xsd:boolean should be one of: 0, 1, true, or false, found " + bool_string);
#else
    return false;
#endif
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

void check_document_type(const std::string &document_content_type)
{
    if (document_content_type != "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" &&
        document_content_type != "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml ")
    {
        throw xlnt::invalid_file(document_content_type);
    }
}

#ifdef THROW_ON_INVALID_XML
#define unexpected_element(element) throw xlnt::exception(element.string());
#else
#define unexpected_element(element) skip_remaining_content(element);
#endif

template<typename T>
bool contains(const std::vector<T> &strings, const T &element)
{
    return std::find(strings.begin(), strings.end(), element) != strings.end();
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


std::vector<relationship> xlsx_consumer::read_relationships(const path &part)
{
    std::vector<xlnt::relationship> relationships;
    if (!archive_->has_file(part.string())) return relationships;

    auto &rels_stream = archive_->open(part.string());
    xml::parser parser(rels_stream, part.string());
    parser_ = &parser;

    xlnt::uri source(part.string());

    static const auto &xmlns = xlnt::constants::get_namespace("relationships");

    expect_start_element(xml::qname(xmlns, "Relationships"), xml::content::complex);

    while (in_element(xml::qname(xmlns, "Relationships")))
    {
        expect_start_element(xml::qname(xmlns, "Relationship"), xml::content::simple);
        relationships.emplace_back(parser.attribute("Id"),
            parser.attribute<xlnt::relationship::type>("Type"), source,
            xlnt::uri(parser.attribute("Target")),
            xlnt::target_mode::internal);
        expect_end_element(xml::qname(xmlns, "Relationship"));
    }

    expect_end_element(xml::qname(xmlns, "Relationships"));
    
    parser_ = nullptr;

    return relationships;
}


void xlsx_consumer::read_part(const std::vector<relationship> &rel_chain)
{
    // ignore namespace declarations except in parts of these types
    const auto using_namespaces = std::vector<relationship::type>
    {
        relationship::type::office_document,
        relationship::type::stylesheet,
        relationship::type::chartsheet,
        relationship::type::dialogsheet,
        relationship::type::worksheet
    };

    auto receive = xml::parser::receive_default;

    if (contains(using_namespaces, rel_chain.back().get_type()))
    {
        receive |= xml::parser::receive_namespace_decls;
    }

    const auto &manifest = target_.get_manifest();
    auto part_path = manifest.canonicalize(rel_chain);
    auto &stream = archive_->open(part_path.string());
    xml::parser parser(stream, part_path.string(), receive);
    parser_ = &parser;

    switch (rel_chain.back().get_type())
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

    case relationship::type::shared_string_table:
        read_shared_string_table();
        break;

    case relationship::type::stylesheet:
        read_stylesheet();
        break;

    case relationship::type::theme:
        read_theme();
        break;

    case relationship::type::chartsheet:
        read_chartsheet(rel_chain.back().get_id());
        break;

    case relationship::type::dialogsheet:
        read_dialogsheet(rel_chain.back().get_id());
        break;

    case relationship::type::worksheet:
        read_worksheet(rel_chain.back().get_id());
        break;

    case relationship::type::thumbnail:
        break;

    case relationship::type::calculation_chain:
        read_calculation_chain();
        break;

    case relationship::type::hyperlink:
        break;

    case relationship::type::comments:
        break;

    case relationship::type::vml_drawing:
        break;

    case relationship::type::unknown:
        break;

    case relationship::type::printer_settings:
        break;

    case relationship::type::custom_property:
        break;

    case relationship::type::drawings:
        break;

    case relationship::type::pivot_table_cache_definition:
        break;

    case relationship::type::pivot_table_cache_records:
        break;

    case relationship::type::query_table:
        break;

    case relationship::type::shared_workbook:
        break;

    case relationship::type::revision_log:
        break;

    case relationship::type::shared_workbook_user_data:
        break;

    case relationship::type::single_cell_table_definitions:
        break;

    case relationship::type::table_definition:
        break;

    case relationship::type::image:
        break;
    }

    parser_ = nullptr;
}

void xlsx_consumer::populate_workbook()
{
    target_.clear();

    read_manifest();
    auto &manifest = target_.get_manifest();

    auto package_parts = std::vector<relationship::type>
    {
        relationship::type::core_properties,
        relationship::type::extended_properties,
        relationship::type::custom_properties,
        relationship::type::office_document,
        relationship::type::connections,
        relationship::type::custom_xml_mappings,
        relationship::type::external_workbook_references,
        relationship::type::metadata,
        relationship::type::pivot_table,
        relationship::type::shared_workbook_revision_headers,
        relationship::type::volatile_dependencies
    };

    auto root_path = path("/");

    for (auto rel_type : package_parts)
    {
        if (manifest.has_relationship(root_path, rel_type))
        {
            read_part({manifest.get_relationship(root_path, rel_type)});
        }
    }

    const auto workbook_rel = manifest.get_relationship(root_path, relationship::type::office_document);

    const auto workbook_parts_first = std::vector<relationship::type>
    {
        relationship::type::shared_string_table,
        relationship::type::stylesheet,
        relationship::type::theme
    };

    for (auto rel_type : workbook_parts_first)
    {
        if (manifest.has_relationship(workbook_rel.get_target().get_path(), rel_type))
        {
            read_part({workbook_rel, manifest.get_relationship(workbook_rel.get_target().get_path(), rel_type)});
        }
    }

    const auto workbook_parts_second = std::vector<relationship::type>
    {
        relationship::type::worksheet,
        relationship::type::chartsheet,
        relationship::type::dialogsheet
    };

    for (auto rel_type : workbook_parts_second)
    {
        if (manifest.has_relationship(workbook_rel.get_target().get_path(), rel_type))
        {
            for (auto rel : manifest.get_relationships(workbook_rel.get_target().get_path(), rel_type))
            {
                read_part({workbook_rel, rel});
            }
        }
    }
}

// Package Parts

void xlsx_consumer::read_manifest()
{
    path package_rels_path("_rels/.rels");

    if (!archive_->has_file(package_rels_path.string()))
    {
        throw invalid_file("missing package rels");
    }

    auto package_rels = read_relationships(package_rels_path);
    auto &manifest = target_.get_manifest();

    static const auto &xmlns = constants::get_namespace("content-types");

    xml::parser parser(archive_->open("[Content_Types].xml"), "[Content_Types].xml");
    parser_ = &parser;

    parser.next_expect(xml::parser::event_type::start_element, xmlns, "Types");
    parser.content(xml::content::complex);

    while (in_element(xml::qname(xmlns, "Types")))
    {
        auto current_element = expect_start_element(xml::content::complex);

        if (current_element == xml::qname(xmlns, "Default"))
        {
            auto extension = parser.attribute("Extension");
            auto content_type = parser.attribute("ContentType");
            manifest.register_default_type(extension, content_type);
        }
        else if (current_element == xml::qname(xmlns, "Override"))
        {
            auto part_name = parser.attribute("PartName");
            auto content_type = parser.attribute("ContentType");
            manifest.register_override_type(path(part_name), content_type);
        }
        else
        {
            unexpected_element(current_element);
        }

        expect_end_element(current_element);
    }

    expect_end_element(xml::qname(xmlns, "Types"));

    for (const auto &package_rel : package_rels)
    {
        if (package_rel.get_type() == relationship::type::office_document)
        {
            auto content_type = manifest.get_content_type(manifest.canonicalize({ package_rel }));
            check_document_type(content_type);
        }

        manifest.register_relationship(uri("/"),
            package_rel.get_type(),
            package_rel.get_target(),
            package_rel.get_target_mode(),
            package_rel.get_id());
    }

    for (const auto &relationship_source_string : archive_->files())
    {
        auto relationship_source = path(relationship_source_string);

        if (relationship_source == path("_rels/.rels") || relationship_source.extension() != "rels") continue;

        path part(relationship_source.parent().parent());
        part = part.append(relationship_source.split_extension().first);
        uri source(part.string());

        path source_directory = part.parent();

        auto part_rels = read_relationships(relationship_source);

        for (const auto &part_rel : part_rels)
        {
            path target_path(source_directory.append(part_rel.get_target().get_path()));
            manifest.register_relationship(source,
                part_rel.get_type(),
                part_rel.get_target(),
                part_rel.get_target_mode(),
                part_rel.get_id());
        }
    }
}

void xlsx_consumer::read_extended_properties()
{
    static const auto &xmlns = constants::get_namespace("extended-properties");
    static const auto &xmlns_vt = constants::get_namespace("vt");

    parser().next_expect(xml::parser::event_type::start_element, xmlns, "Properties");
    parser().content(xml::parser::content_type::complex);

    while (in_element(xml::qname(xmlns, "Properties")))
    {
        auto current_property_element = expect_start_element(xml::content::mixed);

        if (current_property_element == xml::qname(xmlns, "Application"))
        {
            target_.set_application(read_text());
        }
        else if (current_property_element == xml::qname(xmlns, "DocSecurity"))
        {
            target_.set_doc_security(std::stoi(read_text()));
        }
        else if (current_property_element == xml::qname(xmlns, "ScaleCrop"))
        {
            target_.set_scale_crop(is_true(read_text()));
        }
        else if (current_property_element == xml::qname(xmlns, "Company"))
        {
            target_.set_company(read_text());
        }
        else if (current_property_element == xml::qname(xmlns, "SharedDoc"))
        {
            target_.set_shared_doc(is_true(read_text()));
        }
        else if (current_property_element == xml::qname(xmlns, "HyperlinksChanged"))
        {
            target_.set_hyperlinks_changed(is_true(read_text()));
        }
        else if (current_property_element == xml::qname(xmlns, "AppVersion"))
        {
            target_.set_app_version(read_text());
        }
        else if (current_property_element == xml::qname(xmlns, "LinksUpToDate"))
        {
            target_.set_links_up_to_date(is_true(read_text()));
        }
        else if (current_property_element == xml::qname(xmlns, "Template"))
        {
            read_text();
        }
        else if (current_property_element == xml::qname(xmlns, "TotalTime"))
        {
            read_text();
        }
        else if (current_property_element == xml::qname(xmlns, "HeadingPairs"))
        {
            parser().content(xml::parser::content_type::complex);
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
        else if (current_property_element == xml::qname(xmlns, "TitlesOfParts"))
        {
            parser().content(xml::parser::content_type::complex);
            parser().next_expect(xml::parser::event_type::start_element, xmlns_vt, "vector");
            parser().content(xml::parser::content_type::complex);

            auto size = parser().attribute<std::size_t>("size");
            parser().attribute("baseType");

            for (auto i = std::size_t(0); i < size; ++i)
            {
                parser().next_expect(xml::parser::event_type::start_element, xmlns_vt, "lpstr");
                parser().content(xml::parser::content_type::simple);
                parser().next_expect(xml::parser::event_type::characters);
                parser().next_expect(xml::parser::event_type::end_element, xmlns_vt, "lpstr");
            }

            parser().next_expect(xml::parser::event_type::end_element, xmlns_vt, "vector");
        }
        else
        {
            unexpected_element(current_property_element);
        }

        parser().next_expect(xml::parser::event_type::end_element, current_property_element);
    }
    
    expect_end_element(xml::qname(xmlns, "Properties"));
}

void xlsx_consumer::read_core_properties()
{
    static const auto &xmlns_cp = constants::get_namespace("core-properties");
    static const auto &xmlns_dc = constants::get_namespace("dc");
    static const auto &xmlns_dcterms = constants::get_namespace("dcterms");
    static const auto &xmlns_xsi = constants::get_namespace("xsi");

    parser().next_expect(xml::parser::event_type::start_element, xmlns_cp, "coreProperties");
    parser().content(xml::parser::content_type::complex);

    while (in_element(xml::qname(xmlns_cp, "coreProperties")))
    {
        auto property_element = expect_start_element(xml::content::simple);
        auto property_value = std::string();

        if (in_element(property_element))
        {
            property_value = read_text();
        }

        if (property_element == xml::qname(xmlns_dc, "creator"))
        {
            target_.set_creator(property_value);
        }
        else if (property_element == xml::qname(xmlns_cp, "lastModifiedBy"))
        {
            target_.set_last_modified_by(property_value);
        }
        else if (property_element == xml::qname(xmlns_dcterms, "created"))
        {
            skip_attributes({xml::qname(xmlns_xsi, "type")});
            target_.set_created(w3cdtf_to_datetime(property_value));
        }
        else if (property_element == xml::qname(xmlns_dcterms, "modified"))
        {
            skip_attributes({xml::qname(xmlns_xsi, "type")});
            target_.set_modified(w3cdtf_to_datetime(property_value));
        }
        else if (property_element == xml::qname(xmlns_dc, "description"))
        {
            // ignore
        }
        else if (property_element == xml::qname(xmlns_dc, "language"))
        {
            // ignore
        }
        else if (property_element == xml::qname(xmlns_dc, "subject"))
        {
            // ignore
        }
        else if (property_element == xml::qname(xmlns_dc, "title"))
        {
            // ignore
        }
        else if (property_element == xml::qname(xmlns_cp, "revision"))
        {
            // ignore
        }
        else
        {
            unexpected_element(property_element);
        }

        expect_end_element(property_element);
    }

    expect_end_element(xml::qname(xmlns_cp, "coreProperties"));
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
    static const auto &xmlns_s = constants::get_namespace("spreadsheetml");
    static const auto &xmlns_x15 = constants::get_namespace("x15");
    static const auto &xmlns_x15ac = constants::get_namespace("x15ac");

    expect_start_element(xml::qname(xmlns, "workbook"), xml::content::complex);

    if (contains(read_namespaces(), xmlns_x15))
    {
        target_.enable_x15();
    }

    skip_attributes({xml::qname(xmlns_mc, "Ignorable")});

    while (in_element(xml::qname(xmlns, "workbook")))
    {
        auto current_workbook_element = expect_start_element(xml::content::complex);

        if (current_workbook_element == xml::qname(xmlns, "fileVersion"))
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
        }
        else if (current_workbook_element == xml::qname(xmlns_mc, "AlternateContent"))
        {
            if (parser().peek() == xml::parser::event_type::start_namespace_decl)
            {
                parser().next_expect(xml::parser::event_type::start_namespace_decl);
            }

            parser().next_expect(xml::parser::event_type::start_element, xmlns_mc, "Choice");
            parser().content(xml::parser::content_type::complex);
            skip_attributes({"Requires"});
            parser().next_expect(xml::parser::event_type::start_element, xmlns_x15ac, "absPath");
            target_.set_absolute_path(path(parser().attribute("url")));

            if (parser().peek() == xml::parser::event_type::start_namespace_decl)
            {
                parser().next_expect(xml::parser::event_type::start_namespace_decl);
                parser().next_expect(xml::parser::event_type::end_namespace_decl);
            }

            parser().next_expect(xml::parser::event_type::end_element, xmlns_x15ac, "absPath");
            parser().next_expect(xml::parser::event_type::end_element, xmlns_mc, "Choice");
        }
        else if (current_workbook_element == xml::qname(xmlns, "bookViews"))
        {
            while (in_element(xml::qname(xmlns, "bookViews")))
            {
                expect_start_element(xml::qname(xmlns, "workbookView"), xml::content::simple);

                skip_attributes({"activeTab", "firstSheet", "showHorizontalScroll",
                    "showSheetTabs", "showVerticalScroll"});

                workbook_view view;
                view.x_window = string_to_size_t(parser().attribute("xWindow"));
                view.y_window = string_to_size_t(parser().attribute("yWindow"));
                view.window_width = string_to_size_t(parser().attribute("windowWidth"));
                view.window_height = string_to_size_t(parser().attribute("windowHeight"));

                if (parser().attribute_present("tabRatio"))
                {
                    view.tab_ratio = string_to_size_t(parser().attribute("tabRatio"));
                }

                expect_end_element(xml::qname(xmlns, "workbookView"));

                target_.set_view(view);
            }
        }
        else if (current_workbook_element == xml::qname(xmlns, "workbookPr"))
        {
            target_.d_->has_properties_ = true;

            if (parser().attribute_present("date1904"))
            {
                target_.set_base_date(is_true(parser().attribute("date1904"))
                    ? calendar::mac_1904 : calendar::windows_1900);
            }

            skip_attributes({"codeName", "defaultThemeVersion",
                "backupFile", "showObjects", "filterPrivacy", "dateCompatibility"});
        }
        else if (current_workbook_element == xml::qname(xmlns, "sheets"))
        {
            std::size_t index = 0;

            while (in_element(xml::qname(xmlns, "sheets")))
            {
                expect_start_element(xml::qname(xmlns_s, "sheet"), xml::content::simple);

                auto title = parser().attribute("name");
                skip_attributes({"state"});

                sheet_title_index_map_[title] = index++;
                sheet_title_id_map_[title] = string_to_size_t(parser().attribute("sheetId"));
                target_.d_->sheet_title_rel_id_map_[title] = parser().attribute(xml::qname(xmlns_r, "id"));

                expect_end_element(xml::qname(xmlns_s, "sheet"));
            }
        }
        else if (current_workbook_element == xml::qname(xmlns, "calcPr"))
        {
            target_.d_->has_calculation_properties_ = true;
            skip_attributes();
        }
        else if (current_workbook_element == xml::qname(xmlns, "extLst"))
        {
            while (in_element(xml::qname(xmlns, "extLst")))
            {
                parser().next_expect(xml::parser::event_type::start_element, xmlns, "ext");
                parser().content(xml::parser::content_type::complex);
                parser().attribute("uri");
                parser().next_expect(xml::parser::event_type::start_namespace_decl);

                parser().next_expect(xml::parser::event_type::start_element);

                skip_attributes({"stringRefSyntax"});

                if (parser().qname() == xml::qname(xmlns_mx, "ArchID"))
                {
                    target_.d_->has_arch_id_ = true;
                    skip_attributes({"Flags"});
                }
                else if (parser().qname() == xml::qname(xmlns_x15, "workbookPr"))
                {
                    skip_attributes({"chartTrackingRefBase"});
                }

                parser().next_expect(xml::parser::event_type::end_element);

                parser().next_expect(xml::parser::event_type::end_element, xmlns, "ext");
                parser().next_expect(xml::parser::event_type::end_namespace_decl);
            }
        }
        else if (current_workbook_element == xml::qname(xmlns, "workbookProtection"))
        {
            skip_remaining_content(xml::qname(xmlns, "workbookProtection"));
        }
        else
        {
            unexpected_element(current_workbook_element);
        }

        expect_end_element(current_workbook_element);
    }

    expect_end_element(xml::qname(xmlns, "workbook"));
}

// Write Workbook Relationship Target Parts

void xlsx_consumer::read_calculation_chain()
{
}

void xlsx_consumer::read_chartsheet(const std::string & /*title*/)
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

void xlsx_consumer::read_dialogsheet(const std::string & /*title*/)
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
    static const auto &xmlns = constants::get_namespace("spreadsheetml");

    parser().next_expect(xml::parser::event_type::start_element, xmlns, "sst");
    parser().content(xml::content::complex);

    skip_attributes({"count"});

    std::size_t unique_count = 0;

    if (parser().attribute_present("uniqueCount"))
    {
        unique_count = string_to_size_t(parser().attribute("uniqueCount"));
    }

    auto &strings = target_.get_shared_strings();

    while (in_element(xml::qname(xmlns, "sst")))
    {
        expect_start_element(xml::qname(xmlns, "si"), xml::content::complex);
        strings.push_back(read_formatted_text(xmlns));
        expect_end_element(xml::qname(xmlns, "si"));
    }

    expect_end_element(xml::qname(xmlns, "sst"));

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
    static const auto &xmlns = constants::get_namespace("spreadsheetml");
    static const auto &xmlns_mc = constants::get_namespace("mc");
    //static const auto &xmlns_x14 = constants::get_namespace("x14");
    static const auto &xmlns_x14ac = constants::get_namespace("x14ac");
    //static const auto &xmlns_x15 = constants::get_namespace("x15");

    auto &stylesheet = target_.impl().stylesheet_;
    
    //todo: should this really be defined here?
    //todo: maybe xlnt::style and xlnt::format can be used here instead now
    struct formatting_record
    {
        std::pair<class alignment, bool> alignment = {{}, 0};
        std::pair<std::size_t, bool> border_id = {0, false};
        std::pair<std::size_t, bool> fill_id = {0, false};
        std::pair<std::size_t, bool> font_id = {0, false};
        std::pair<std::size_t, bool> number_format_id = {0, false};
        std::pair<class protection, bool> protection = {{}, false};
        std::pair<std::size_t, bool> style_id = {0, false};
    };

    struct style_data
    {
        std::string name;
        std::size_t record_id;
        std::pair<std::size_t, bool> builtin_id;
        std::pair<std::size_t, bool> i_level;
        std::pair<bool, bool> hidden;
        bool custom_builtin;
    };

    std::vector<style_data> style_datas;
    std::vector<formatting_record> style_records;
    std::vector<formatting_record> format_records;

    parser().next_expect(xml::parser::event_type::start_element, xmlns, "styleSheet");
    parser().content(xml::parser::content_type::complex);

    if (contains(read_namespaces(), xmlns_x14ac))
    {
        //todo: not enable_x14ac?
        target_.enable_x15();
    }

    skip_attributes({xml::qname(xmlns_mc, "Ignorable")});

    while (in_element(xml::qname(xmlns, "styleSheet")))
    {
        auto current_style_element = expect_start_element(xml::content::complex);

        if (current_style_element == xml::qname(xmlns, "borders"))
        {
            stylesheet.borders.clear();

            auto count = parser().attribute<std::size_t>("count");

            while (in_element(xml::qname(xmlns, "borders")))
            {
                stylesheet.borders.push_back(xlnt::border());
                auto &border = stylesheet.borders.back();

                expect_start_element(xml::qname(xmlns, "border"), xml::content::complex);

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

                while (in_element(xml::qname(xmlns, "border")))
                {
                    auto current_side_element = expect_start_element(xml::content::complex);

                    xlnt::border::border_property side;

                    if (parser().attribute_present("style"))
                    {
                        side.style(parser().attribute<xlnt::border_style>("style"));
                    }

                    if (in_element(current_side_element))
                    {
                        expect_start_element(xml::qname(xmlns, "color"), xml::content::complex);
                        side.color(read_color(parser()));
                        expect_end_element(xml::qname(xmlns, "color"));
                    }

                    expect_end_element(current_side_element);

                    auto side_type = xml::value_traits<xlnt::border_side>::parse(current_side_element.name(), parser());
                    border.side(side_type, side);
                }

                expect_end_element(xml::qname(xmlns, "border"));
            }

            if (count != stylesheet.borders.size())
            {
                throw xlnt::exception("border counts don't match");
            }
        }
        else if (current_style_element == xml::qname(xmlns, "fills"))
        {
            stylesheet.fills.clear();

            auto count = parser().attribute<std::size_t>("count");

            while (in_element(xml::qname(xmlns, "fills")))
            {
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

                        while (in_element(xml::qname(xmlns, "patternType")))
                        {
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

                    while (in_element(xml::qname(xmlns, "gradientFill")))
                    {
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
        else if (current_style_element == xml::qname(xmlns, "fonts"))
        {
            stylesheet.fonts.clear();

            auto count = parser().attribute<std::size_t>("count");
            skip_attributes({xml::qname(xmlns_x14ac, "knownFonts")});

            while (in_element(xml::qname(xmlns, "fonts")))
            {
                stylesheet.fonts.push_back(xlnt::font());
                auto &new_font = stylesheet.fonts.back();

                parser().next_expect(xml::parser::event_type::start_element, xmlns, "font");
                parser().content(xml::parser::content_type::complex);

                while (in_element(xml::qname(xmlns, "font")))
                {
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
        else if (current_style_element == xml::qname(xmlns, "numFmts"))
        {
            stylesheet.number_formats.clear();

            auto count = parser().attribute<std::size_t>("count");

            while (in_element(xml::qname(xmlns, "numFmts")))
            {
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
        else if (current_style_element == xml::qname(xmlns, "colors"))
        {
            skip_remaining_content(xml::qname(xmlns, "colors"));
        }
        else if (current_style_element == xml::qname(xmlns, "cellStyles"))
        {
            auto count = parser().attribute<std::size_t>("count");

            while (in_element(xml::qname(xmlns, "cellStyles")))
            {
                auto &data = *style_datas.emplace(style_datas.end());

                parser().next_expect(xml::parser::event_type::start_element, xmlns, "cellStyle");

                data.name = parser().attribute("name");
                data.record_id = parser().attribute<std::size_t>("xfId");

                if (parser().attribute_present("builtinId"))
                {
                    data.builtin_id = {parser().attribute<std::size_t>("builtinId"), true};
                }

                if (parser().attribute_present("iLevel"))
                {
                    data.i_level = {parser().attribute<std::size_t>("iLevel"), true};
                }

                if (parser().attribute_present("hidden"))
                {
                    data.hidden = {is_true(parser().attribute("hidden")), true};
                }

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
        else if (current_style_element == xml::qname(xmlns, "cellStyleXfs") ||
                 current_style_element == xml::qname(xmlns, "cellXfs"))
        {
            auto in_style_records = current_style_element.name() == "cellStyleXfs";
            auto count = parser().attribute<std::size_t>("count");

            while (in_element(current_style_element))
            {
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
                record.border_id = {border_index, border_applied};

                auto fill_applied = parser().attribute_present("applyFill")
                    && is_true(parser().attribute("applyFill"));
                auto fill_index = parser().attribute_present("fillId")
                    ? string_to_size_t(parser().attribute("fillId")) : 0;
                record.fill_id = {fill_index, fill_applied};

                auto font_applied = parser().attribute_present("applyFont")
                    && is_true(parser().attribute("applyFont"));
                auto font_index = parser().attribute_present("fontId")
                    ? string_to_size_t(parser().attribute("fontId")) : 0;
                record.font_id = {font_index, font_applied};

                auto number_format_applied = parser().attribute_present("applyNumberFormat")
                    && is_true(parser().attribute("applyNumberFormat"));
                auto number_format_id = parser().attribute_present("numFmtId")
                    ? string_to_size_t(parser().attribute("numFmtId")) : 0;
                record.number_format_id = {number_format_id, number_format_applied};

                auto apply_protection_present = parser().attribute_present("applyProtection");
                auto protection_applied = apply_protection_present
                    && is_true(parser().attribute("applyProtection"));
                record.protection.second = protection_applied;

                if (parser().attribute_present("xfId") && parser().name() == "cellXfs")
                {
                    record.style_id = {parser().attribute<std::size_t>("xfId"), true};
                }

                while (in_element(xml::qname(xmlns, "xf")))
                {
                    parser().next_expect(xml::parser::event_type::start_element);

                    if (parser().qname() == xml::qname(xmlns, "alignment"))
                    {
                        if (parser().attribute_present("wrapText"))
                        {
                            auto value = is_true(parser().attribute("wrapText"));
                            record.alignment.first.wrap(value);
                        }

                        if (parser().attribute_present("shrinkToFit"))
                        {
                            auto value = is_true(parser().attribute("shrinkToFit"));
                            record.alignment.first.shrink(value);
                        }

                        if (parser().attribute_present("indent"))
                        {
                            auto value = parser().attribute<int>("indent");
                            record.alignment.first.indent(value);
                        }

                        if (parser().attribute_present("textRotation"))
                        {
                            auto value = parser().attribute<int>("textRotation");
                            record.alignment.first.rotation(value);
                        }

                        if (parser().attribute_present("vertical"))
                        {
                            auto value = parser().attribute<xlnt::vertical_alignment>("vertical");
                            record.alignment.first.vertical(value);
                        }

                        if (parser().attribute_present("horizontal"))
                        {
                            auto value = parser().attribute<xlnt::horizontal_alignment>("horizontal");
                            record.alignment.first.horizontal(value);
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

            if ((in_style_records && count != style_records.size()) ||
                (!in_style_records && count != format_records.size()))
            {
                throw xlnt::exception("counts don't match");
            }
        }
        else if (current_style_element == xml::qname(xmlns, "dxfs"))
        {
            auto count = parser().attribute<std::size_t>("count");
            std::size_t processed = 0;

            while (in_element(xml::qname(xmlns, "dxfs")))
            {
                auto current_element = expect_start_element(xml::content::complex);
                skip_remaining_content(current_element);
		
                ++processed;
            }

            if (count != processed)
            {
                throw xlnt::exception("counts don't match");
            }
        }
        else if (current_style_element == xml::qname(xmlns, "tableStyles"))
        {
            skip_attributes(std::vector<std::string>{"defaultTableStyle", "defaultPivotStyle"});

            auto count = parser().attribute<std::size_t>("count");
            std::size_t processed = 0;

            while (in_element(xml::qname(xmlns, "tableStyles")))
            {
                auto current_element = expect_start_element(xml::content::complex);
                skip_remaining_content(current_element);

                ++processed;
            }

            if (count != processed)
            {
                throw xlnt::exception("counts don't match");
            }
        }
        else if (current_style_element == xml::qname(xmlns, "extLst"))
        {
            while (in_element(xml::qname(xmlns, "extLst")))
            {
                parser().next_expect(xml::parser::event_type::start_element, xmlns, "ext");
                parser().content(xml::parser::content_type::complex);
                skip_attributes({"uri"});
                parser().next_expect(xml::parser::event_type::start_namespace_decl);

                parser().next_expect(xml::parser::event_type::start_element);
                skip_attributes();
                parser().next_expect(xml::parser::event_type::end_element);

                parser().next_expect(xml::parser::event_type::end_element, xmlns, "ext");
                parser().next_expect(xml::parser::event_type::end_namespace_decl);
            }
        }
        else if (current_style_element == xml::qname(xmlns, "indexedColors"))
        {
            skip_remaining_content(current_style_element);
        }
        else
        {
            unexpected_element(current_style_element);
        }

        expect_end_element(current_style_element);
    }

    expect_end_element(xml::qname(xmlns, "styleSheet"));

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

        if (style_data_iter->builtin_id.second)
        {
            new_style.builtin_id(style_data_iter->builtin_id.first);
        }

        if (style_data_iter->hidden.second)
        {
            new_style.hidden(style_data_iter->hidden.first);
        }

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

// CT_Worksheet
void xlsx_consumer::read_worksheet(const std::string &rel_id)
{
    static const auto &xmlns = constants::get_namespace("spreadsheetml");
    static const auto &xmlns_mc = constants::get_namespace("mc");
    static const auto &xmlns_x14ac = constants::get_namespace("x14ac");

    auto title = std::find_if(target_.d_->sheet_title_rel_id_map_.begin(),
        target_.d_->sheet_title_rel_id_map_.end(),
        [&](const std::pair<std::string, std::string> &p) { return p.second == rel_id; })->first;

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

    // begin CT_Worksheet

    parser().next_expect(xml::parser::event_type::start_element, xmlns, "worksheet");
    parser().content(xml::parser::content_type::complex);

    if (contains(read_namespaces(), xmlns_x14ac))
    {
        ws.enable_x14ac();
    }

    skip_attributes({xml::qname(xmlns_mc, "Ignorable")});

    xlnt::range_reference full_range;
    const auto &shared_strings = target_.get_shared_strings();

    while (in_element(xml::qname(xmlns, "worksheet")))
    {
        auto current_worksheet_element = expect_start_element(xml::content::complex);

        if (current_worksheet_element == xml::qname(xmlns, "sheetPr")) // CT_SheetPr 0-1
        {
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == xml::qname(xmlns, "dimension")) // CT_SheetDimension 0-1
        {
            full_range = xlnt::range_reference(parser().attribute("ref"));
            ws.d_->has_dimension_ = true;
        }
        else if (current_worksheet_element == xml::qname(xmlns, "sheetViews")) // CT_SheetViews 0-1
        {
            ws.d_->has_view_ = true;
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == xml::qname(xmlns, "sheetFormatPr")) // CT_SheetFormatPr 0-1
        {
            ws.d_->has_format_properties_ = true;
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == xml::qname(xmlns, "cols")) // CT_Cols 0+
        {
            while (in_element(xml::qname(xmlns, "cols")))
            {
                parser().next_expect(xml::parser::event_type::start_element, xmlns, "col");

                skip_attributes({"bestFit", "collapsed", "hidden", "outlineLevel"});

                auto min = static_cast<column_t::index_t>(std::stoull(parser().attribute("min")));
                auto max = static_cast<column_t::index_t>(std::stoull(parser().attribute("max")));

                auto width = std::stold(parser().attribute("width"));
                auto column_style = parser().attribute_present("style")
                    ? parser().attribute<std::size_t>("style") : static_cast<std::size_t>(0);
                auto custom = parser().attribute_present("customWidth")
                    ? is_true(parser().attribute("customWidth")) : false;

                parser().next_expect(xml::parser::event_type::end_element, xmlns, "col");

                for (auto column = min; column <= max; column++)
                {
                    column_properties props;

                    props.width = width;
                    props.style = column_style;
                    props.custom = custom;

                    ws.add_column_properties(column, props);
                }
            }
        }
        else if (current_worksheet_element == xml::qname(xmlns, "sheetData")) // CT_SheetData 1
        {
            while (in_element(xml::qname(xmlns, "sheetData")))
            {
                parser().next_expect(xml::parser::event_type::start_element, xmlns, "row"); // CT_Row
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

                skip_attributes({xml::qname(xmlns_x14ac, "dyDescent")});
                skip_attributes({"customFormat", "s", "customFont", "hidden",
                    "outlineLevel", "collapsed", "thickTop", "thickBot", "ph"});

                while (in_element(xml::qname(xmlns, "row")))
                {
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

                    while (in_element(xml::qname(xmlns, "c")))
                    {
                        auto current_element = expect_start_element(xml::content::mixed);
                        
                        if (current_element == xml::qname(xmlns, "v")) // s:ST_Xstring
                        {
                            has_value = true;
                            value_string = read_text();
                        }
                        else if (current_element == xml::qname(xmlns, "f")) // CT_CellFormula
                        {
                            has_formula = true;

                            if (parser().attribute_present("t"))
                            {
                                has_shared_formula = parser().attribute("t") == "shared";
                            }

                            skip_attributes({"aca", "ref", "dt2D", "dtr", "del1", "del2", "r1", "r2", "ca", "si", "bx"});

                            formula_value_string = read_text();
                        }
                        else if (current_element == xml::qname(xmlns, "is")) // CT_Rst
                        {
                            parser().next_expect(xml::parser::event_type::start_element, xmlns, "t");
                            value_string = read_text();
                            parser().next_expect(xml::parser::event_type::end_element, xmlns, "t");
                        }
                        else
                        {
                            unexpected_element(current_element);
                        }

                        parser().next_expect(xml::parser::event_type::end_element, current_element);
                    }

                    parser().next_expect(xml::parser::event_type::end_element, xmlns, "c");

                    if (has_formula && !has_shared_formula && !ws.get_workbook().get_data_only())
                    {
                        cell.set_formula(formula_value_string);
                    }

                    if (has_type && (type == "inlineStr" || type == "str"))
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
                }

                expect_end_element(xml::qname(xmlns, "row"));
            }
        }
        else if (current_worksheet_element == xml::qname(xmlns, "sheetCalcPr")) // CT_SheetCalcPr 0-1
        {
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == xml::qname(xmlns, "sheetProtection")) // CT_SheetProtection 0-1
        {
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == xml::qname(xmlns, "protectedRanges")) // CT_ProtectedRanges 0-1
        {
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == xml::qname(xmlns, "scenarios")) // CT_Scenarios 0-1
        {
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == xml::qname(xmlns, "autoFilter")) // CT_AutoFilter 0-1
        {
            ws.auto_filter(xlnt::range_reference(parser().attribute("ref")));
        }
        else if (current_worksheet_element == xml::qname(xmlns, "sortState")) // CT_SortState 0-1
        {
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == xml::qname(xmlns, "dataConsolidate")) // CT_DataConsolidate 0-1
        {
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == xml::qname(xmlns, "customSheetViews")) // CT_CustomSheetViews 0-1
        {
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == xml::qname(xmlns, "mergeCells")) // CT_MergeCells 0-1
        {
            auto count = std::stoull(parser().attribute("count"));

            while (in_element(xml::qname(xmlns, "mergeCells")))
            {
                parser().next_expect(xml::parser::event_type::start_element, xmlns, "mergeCell");
                ws.merge_cells(range_reference(parser().attribute("ref")));
                parser().next_expect(xml::parser::event_type::end_element, xmlns, "mergeCell");

                count--;
            }

            if (count != 0)
            {
                throw invalid_file("sizes don't match");
            }
        }
        else if (current_worksheet_element == xml::qname(xmlns, "phoneticPr")) // CT_PhoneticPr 0-1
        {
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == xml::qname(xmlns, "conditionalFormatting")) // CT_ConditionalFormatting 0+
        {
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == xml::qname(xmlns, "dataValidations")) // CT_DataValidations 0-1
        {
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == xml::qname(xmlns, "hyperlinks")) // CT_Hyperlinks 0-1
        {
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == xml::qname(xmlns, "printOptions")) // CT_PrintOptions 0-1
        {
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == xml::qname(xmlns, "pageMargins")) // CT_PageMargins 0-1
        {
            page_margins margins;

            margins.set_top(parser().attribute<double>("top"));
            margins.set_bottom(parser().attribute<double>("bottom"));
            margins.set_left(parser().attribute<double>("left"));
            margins.set_right(parser().attribute<double>("right"));
            margins.set_header(parser().attribute<double>("header"));
            margins.set_footer(parser().attribute<double>("footer"));

            ws.set_page_margins(margins);
        }
        else if (current_worksheet_element == xml::qname(xmlns, "pageSetup")) // CT_PageSetup 0-1
        {
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == xml::qname(xmlns, "headerFooter")) // CT_HeaderFooter 0-1
        {
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == xml::qname(xmlns, "rowBreaks")) // CT_PageBreak 0-1
        {
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == xml::qname(xmlns, "colBreaks")) // CT_PageBreak 0-1
        {
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == xml::qname(xmlns, "customProperties")) // CT_CustomProperties 0-1
        {
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == xml::qname(xmlns, "cellWatches")) // CT_CellWatches 0-1
        {
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == xml::qname(xmlns, "ignoredErrors")) // CT_IgnoredErrors 0-1
        {
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == xml::qname(xmlns, "smartTags")) // CT_SmartTags 0-1
        {
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == xml::qname(xmlns, "drawing")) // CT_Drawing 0-1
        {
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == xml::qname(xmlns, "legacyDrawing"))
        {
            skip_remaining_content(current_worksheet_element);
        }
        else
        {
            unexpected_element(current_worksheet_element);
        }

        parser().next_expect(xml::parser::event_type::end_element, current_worksheet_element);
    }

    parser().next_expect(xml::parser::event_type::end_element, xmlns, "worksheet");

    auto &manifest = target_.get_manifest();
    const auto workbook_rel = manifest.get_relationship(path("/"), relationship::type::office_document);
    const auto sheet_rel = manifest.get_relationship(workbook_rel.get_target().get_path(), rel_id);
    path sheet_path(sheet_rel.get_source().get_path().parent().append(sheet_rel.get_target().get_path()));

    if (manifest.has_relationship(sheet_path, xlnt::relationship::type::comments))
    {
        auto comments_part = manifest.canonicalize(
            {workbook_rel, sheet_rel, manifest.get_relationship(sheet_path, xlnt::relationship::type::comments)});

        auto receive = xml::parser::receive_default;
        xml::parser parser(archive_->open(comments_part.string()), comments_part.string(), receive);
        parser_ = &parser;

        read_comments(ws);

        if (manifest.has_relationship(sheet_path, xlnt::relationship::type::vml_drawing))
        {
            auto vml_drawings_part = manifest.canonicalize({workbook_rel, sheet_rel,
                manifest.get_relationship(sheet_path, xlnt::relationship::type::vml_drawing)});

            xml::parser vml_parser(archive_->open(vml_drawings_part.string()), vml_drawings_part.string(), receive);
            parser_ = &vml_parser;

            read_vml_drawings(ws);
        }
    }
}

// Sheet Relationship Target Parts

void xlsx_consumer::read_vml_drawings(worksheet/*ws*/)
{
}

void xlsx_consumer::read_comments(worksheet ws)
{
    static const auto &xmlns = xlnt::constants::get_namespace("spreadsheetml");

    std::vector<std::string> authors;

    expect_start_element(xml::qname(xmlns, "comments"), xml::content::complex);
    expect_start_element(xml::qname(xmlns, "authors"), xml::content::complex);

    while (in_element(xml::qname(xmlns, "authors")))
    {
        expect_start_element(xml::qname(xmlns, "author"), xml::content::simple);
        authors.push_back(read_text());
        expect_end_element(xml::qname(xmlns, "author"));
    }

    expect_end_element(xml::qname(xmlns, "authors"));
    expect_start_element(xml::qname(xmlns, "commentList"), xml::content::complex);

    while (in_element(xml::qname(xmlns, "commentList")))
    {
        expect_start_element(xml::qname(xmlns, "comment"), xml::content::complex);

        auto cell_ref = parser().attribute("ref");
        auto author_id = parser().attribute<std::size_t>("authorId");

        expect_start_element(xml::qname(xmlns, "text"), xml::content::complex);

        ws.get_cell(cell_ref).comment(comment(read_formatted_text(xmlns), authors.at(author_id)));

        expect_end_element(xml::qname(xmlns, "text"));
        expect_end_element(xml::qname(xmlns, "comment"));
    }

    expect_end_element(xml::qname(xmlns, "commentList"));
    expect_end_element(xml::qname(xmlns, "comments"));
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

std::string xlsx_consumer::read_text()
{
    auto text = std::string();

    while (parser().peek() == xml::parser::event_type::characters)
    {
        parser().next_expect(xml::parser::event_type::characters);
        text.append(parser().value());
    }

    return text;
}

void xlsx_consumer::skip_attributes(const std::vector<std::string> &names)
{
    for (const auto &name : names)
    {
        if (parser().attribute_present(name))
        {
            parser().attribute(name);
        }
    }
}

void xlsx_consumer::skip_attributes(const std::vector<xml::qname> &names)
{
    for (const auto &name : names)
    {
        if (parser().attribute_present(name))
        {
            parser().attribute(name);
        }
    }
}

void xlsx_consumer::skip_attributes()
{
    parser().attribute_map();
}

void xlsx_consumer::skip_remaining_content(const xml::qname &name)
{
    // start by assuming we've already parsed the opening tag

    // skip attributes on the opening tag
    skip_attributes();
    read_text();

    // continue until the closing tag is reached
    while (in_element(name))
    {
        auto current_element = expect_start_element(xml::content::mixed);
        skip_remaining_content(current_element);
        expect_end_element(current_element);
    }
}

std::vector<std::string> xlsx_consumer::read_namespaces()
{
    std::vector<std::string> namespaces;

    while (parser().peek() == xml::parser::event_type::start_namespace_decl)
    {
        parser().next_expect(xml::parser::event_type::start_namespace_decl);
	namespaces.push_back(parser().namespace_());
    }

    return namespaces;
}

bool xlsx_consumer::in_element(const xml::qname &name)
{
    if (parser().peek() == xml::parser::event_type::end_element)
    {
        return false;
    }

    return true;
}

xml::qname xlsx_consumer::expect_start_element(xml::content content)
{
    parser().next_expect(xml::parser::event_type::start_element);
    parser().content(content);

    return parser().qname();
}

void xlsx_consumer::expect_start_element(const xml::qname &name, xml::content content)
{
    parser().next_expect(xml::parser::event_type::start_element, name);
    parser().content(content);
}

void xlsx_consumer::expect_end_element(const xml::qname &name)
{
    parser().next_expect(xml::parser::event_type::end_element, name);

    while (parser().peek() == xml::parser::event_type::end_namespace_decl)
    {
        parser().next_expect(xml::parser::event_type::end_namespace_decl);
    }
}

formatted_text xlsx_consumer::read_formatted_text(const std::string &xmlns)
{
    formatted_text t;

    auto text_element = expect_start_element(xml::content::mixed);
    skip_attributes();

    if (text_element == xml::qname(xmlns, "t"))
    {
        t.plain_text(read_text());
    }
    else if (text_element == xml::qname(xmlns, "r"))
    {
        parser().content(xml::content::complex);

        text_run run;

        while (in_element(xml::qname(xmlns, "r")))
        {
            auto run_element = expect_start_element(xml::content::mixed);
            
            if (run_element == xml::qname(xmlns, "rPr"))
            {
                parser().content(xml::content::complex);

                while (in_element(xml::qname(xmlns, "rPr")))
                {
                    auto current_run_property_element = expect_start_element(xml::content::simple);

                    if (current_run_property_element == xml::qname(xmlns, "sz"))
                    {
                        run.set_size(string_to_size_t(parser().attribute("val")));
                    }
                    else if (current_run_property_element == xml::qname(xmlns, "rFont"))
                    {
                        run.set_font(parser().attribute("val"));
                    }
                    else if (current_run_property_element == xml::qname(xmlns, "color"))
                    {
                        run.set_color(read_color(parser()));
                    }
                    else if (current_run_property_element == xml::qname(xmlns, "family"))
                    {
                        run.set_family(string_to_size_t(parser().attribute("val")));
                    }
                    else if (current_run_property_element == xml::qname(xmlns, "scheme"))
                    {
                        run.set_scheme(parser().attribute("val"));
                    }

                    expect_end_element(current_run_property_element);
                }
            }
            else if (run_element == xml::qname(xmlns, "t"))
            {
                run.set_string(read_text());
            }
            
            expect_end_element(run_element);
        }
        
        t.add_run(run);
    }
    else
    {
        unexpected_element(text_element);
    }

    expect_end_element(text_element);

    return t;
}

} // namespace detail
} // namepsace xlnt
