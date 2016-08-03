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
#include <algorithm>
#include <iterator>
#include <pugixml.hpp>

#include <detail/constants.hpp>
#include <detail/excel_serializer.hpp>
#include <detail/manifest_serializer.hpp>
#include <detail/relationship_serializer.hpp>
#include <detail/shared_strings_serializer.hpp>
#include <detail/style_serializer.hpp>
#include <detail/stylesheet.hpp>
#include <detail/theme_serializer.hpp>
#include <detail/workbook_impl.hpp>
#include <detail/workbook_serializer.hpp>
#include <detail/worksheet_serializer.hpp>
#include <xlnt/cell/text.hpp>
#include <xlnt/packaging/manifest.hpp>
#include <xlnt/styles/format.hpp>
#include <xlnt/styles/style.hpp>
#include <xlnt/utils/exceptions.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/workbook/worksheet_iterator.hpp>
#include <xlnt/worksheet/range_iterator.hpp>
#include <xlnt/worksheet/worksheet.hpp>

namespace {

std::string::size_type find_string_in_string(const std::string &string, const std::string &substring)
{
    std::string::size_type possible_match_index = string.find(substring.at(0));

    while (possible_match_index != std::string::npos)
    {
        if (string.substr(possible_match_index, substring.size()) == substring)
        {
            return possible_match_index;
        }

        possible_match_index = string.find(substring.at(0), possible_match_index + 1);
    }

    return possible_match_index;
}

bool load_workbook(xlnt::zip_file &archive, bool guess_types, bool data_only, xlnt::workbook &wb, xlnt::detail::stylesheet &stylesheet)
{
    wb.clear();

    if(!archive.has_file(xlnt::constants::part_content_types()))
    {
        throw xlnt::invalid_file("missing [Content Types].xml");
    }
    
    xlnt::manifest_serializer ms(wb.get_manifest());
	pugi::xml_document manifest_xml;
    manifest_xml.load(archive.read(xlnt::constants::part_content_types()).c_str());
    ms.read_manifest(manifest_xml);

    if (ms.determine_document_type() != "excel")
    {
        throw xlnt::invalid_file("package is not an OOXML SpreadsheetML");
    }

    if(archive.has_file(xlnt::constants::part_core()))
    {
        xlnt::workbook_serializer workbook_serializer_(wb);
		pugi::xml_document core_properties_xml;
		core_properties_xml.load(archive.read(xlnt::constants::part_core()).c_str());
        workbook_serializer_.read_properties_core(core_properties_xml);
    }
    
    if(archive.has_file(xlnt::constants::part_app()))
    {
        xlnt::workbook_serializer workbook_serializer_(wb);
		pugi::xml_document app_properties_xml;
		app_properties_xml.load(archive.read(xlnt::constants::part_app()).c_str());
        workbook_serializer_.read_properties_app(app_properties_xml);
    }

    xlnt::relationship_serializer relationship_serializer_(archive);
    auto root_relationships = relationship_serializer_.read_relationships("");
    
    for (const auto &relationship : root_relationships)
    {
        wb.create_root_relationship(relationship.get_id(), relationship.get_target_uri(), relationship.get_type());
    }
    
    auto workbook_relationships = relationship_serializer_.read_relationships(xlnt::constants::part_workbook().to_string('/'));

    for (const auto &relationship : workbook_relationships)
    {
        wb.create_relationship(relationship.get_id(), relationship.get_target_uri(), relationship.get_type());
    }

	pugi::xml_document xml;
    xml.load(archive.read(xlnt::constants::part_workbook()).c_str());

    auto root_node = xml.child("workbook");
    auto workbook_pr_node = root_node.child("workbookPr");

	if (workbook_pr_node.attribute("date1904"))
	{
		std::string value = workbook_pr_node.attribute("date1904").value();

		if (value == "1" || value == "true")
		{
			wb.set_base_date(xlnt::calendar::mac_1904);
		}
	}

    if(archive.has_file(xlnt::constants::part_shared_strings()))
    {
        std::vector<xlnt::text> shared_strings;
		pugi::xml_document shared_strings_xml;
		shared_strings_xml.load(archive.read(xlnt::constants::part_shared_strings()).c_str());
		xlnt::shared_strings_serializer::read_shared_strings(shared_strings_xml, shared_strings);

        for (auto &shared_string : shared_strings)
        {
            wb.add_shared_string(shared_string, true);
        }
    }

    xlnt::style_serializer style_serializer(stylesheet);
	pugi::xml_document style_xml;
	style_xml.load(archive.read(xlnt::constants::part_styles()).c_str());
    style_serializer.read_stylesheet(style_xml);

    for (auto sheet_node : root_node.child("sheets").children())
    {
        auto rel = wb.get_relationship(sheet_node.attribute("r:id").value());

        // TODO impelement chartsheets
        if(rel.get_type() == xlnt::relationship::type::chartsheet)
        {
            continue;
        }

        auto ws = wb.create_sheet_with_rel(sheet_node.attribute("name").value(), rel);
        ws.set_id(static_cast<std::size_t>(sheet_node.attribute("sheetId").as_ullong()));
        xlnt::worksheet_serializer worksheet_serializer(ws);
		pugi::xml_document worksheet_xml;
		auto target_uri = xlnt::constants::package_xl().append(rel.get_target_uri());
		worksheet_xml.load(archive.read(target_uri).c_str());

        worksheet_serializer.read_worksheet(worksheet_xml, stylesheet);
    }

    if (archive.has_file(xlnt::path("docProps/thumbnail.jpeg")))
    {
        auto thumbnail_data = archive.read(xlnt::path("docProps/thumbnail.jpeg"));
        wb.set_thumbnail(std::vector<std::uint8_t>(thumbnail_data.begin(), thumbnail_data.end()));
    }

    return true;
}

} // namespace

namespace xlnt {

const std::string excel_serializer::central_directory_signature()
{
    return "\x50\x4b\x05\x06";
}

std::string excel_serializer::repair_central_directory(const std::string &original)
{
    auto pos = find_string_in_string(original, central_directory_signature());

    if (pos != std::string::npos)
    {
        return original.substr(0, pos + 22);
    }

    return original;
}

bool excel_serializer::load_stream_workbook(std::istream &stream, bool guess_types, bool data_only)
{
    std::vector<std::uint8_t> bytes;
    
    //TODO: inefficient?
    while (stream.good())
    {
        bytes.push_back(static_cast<std::uint8_t>(stream.get()));
    }
    
    return load_virtual_workbook(bytes, guess_types, data_only);
}

bool excel_serializer::load_workbook(const path &filename, bool guess_types, bool data_only)
{
    try
    {
        archive_.load(filename);
    }
    catch (std::runtime_error)
    {
        throw invalid_file(filename.to_string());
    }

    return ::load_workbook(archive_, guess_types, data_only, workbook_, get_stylesheet());
}

bool excel_serializer::load_virtual_workbook(const std::vector<std::uint8_t> &bytes, bool guess_types, bool data_only)
{
    archive_.load(bytes);

    return ::load_workbook(archive_, guess_types, data_only, workbook_, get_stylesheet());
}

excel_serializer::excel_serializer(workbook &wb) : workbook_(wb)
{
}

void excel_serializer::write_data(bool /*as_template*/)
{
    relationship_serializer relationship_serializer_(archive_);
    relationship_serializer_.write_relationships(workbook_.get_root_relationships(), "");
    relationship_serializer_.write_relationships(workbook_.get_relationships(), constants::part_workbook().to_string('/'));

    pugi::xml_document properties_app_xml;
    workbook_serializer workbook_serializer_(workbook_);
    workbook_serializer_.write_properties_app(properties_app_xml);

    {
        std::ostringstream ss;
        properties_app_xml.save(ss);
        archive_.write_string(ss.str(), constants::part_app());
    }

    pugi::xml_document properties_core_xml;
    workbook_serializer_.write_properties_core(properties_core_xml);

    {
        std::ostringstream ss;
        properties_core_xml.save(ss);
        archive_.write_string(ss.str(), constants::part_core());
    }

    pugi::xml_document theme_xml;
    theme_serializer theme_serializer_;
    theme_serializer_.write_theme(workbook_.get_theme(), theme_xml);

    {
        std::ostringstream ss;
        theme_xml.save(ss);
        archive_.write_string(ss.str(), constants::part_theme());
    }

    if (!workbook_.get_shared_strings().empty())
    {
        const auto &strings = workbook_.get_shared_strings();
        pugi::xml_document shared_strings_xml;
        shared_strings_serializer::write_shared_strings(strings, shared_strings_xml);

        std::ostringstream ss;
        shared_strings_xml.save(ss);
        archive_.write_string(ss.str(), constants::part_shared_strings());
    }

    pugi::xml_document workbook_xml;
    workbook_serializer_.write_workbook(workbook_xml);

    {
        std::ostringstream ss;
        workbook_xml.save(ss);
        archive_.write_string(ss.str(), constants::part_workbook());
    }

    style_serializer style_serializer(workbook_.d_->stylesheet_);
    pugi::xml_document style_xml;
    style_serializer.write_stylesheet(style_xml);

    {
        std::ostringstream ss;
        style_xml.save(ss);
        archive_.write_string(ss.str(), constants::part_styles());
    }

    manifest_serializer manifest_serializer_(workbook_.get_manifest());
    pugi::xml_document manifest_xml;
    manifest_serializer_.write_manifest(manifest_xml);

    {
        std::ostringstream ss;
        manifest_xml.save(ss);
        archive_.write_string(ss.str(), constants::part_content_types());
    }

    write_worksheets();

    if(!workbook_.get_thumbnail().empty())
    {
        const auto &thumbnail = workbook_.get_thumbnail();
        archive_.write_string(std::string(thumbnail.begin(), thumbnail.end()), path("docProps/thumbnail.jpeg"));
    }
}

void excel_serializer::write_worksheets()
{
    std::size_t index = 1;

    for (const auto ws : workbook_)
    {
        auto target = "worksheets/sheet" + std::to_string(index++) + ".xml";

        for (const auto &rel : workbook_.get_relationships())
        {
            if (rel.get_target_uri() != target) continue;

            worksheet_serializer serializer_(ws);
			path ws_path(rel.get_target_uri().substr(0, 3) != "xl/" ? "xl/" : "", '/');
			ws_path.append(rel.get_target_uri());
            std::ostringstream ss;
            pugi::xml_document worksheet_xml;
            serializer_.write_worksheet(worksheet_xml);
            worksheet_xml.save(ss);
            archive_.write_string(ss.str(), ws_path);

            break;
        }
    }
}

void excel_serializer::write_external_links()
{
}

bool excel_serializer::save_stream_workbook(std::ostream &stream, bool as_template)
{
    write_data(as_template);
    archive_.save(stream);

    return true;
}

bool excel_serializer::save_workbook(const path &filename, bool as_template)
{
    write_data(as_template);
    archive_.save(filename);

    return true;
}

bool excel_serializer::save_virtual_workbook(std::vector<std::uint8_t> &bytes, bool as_template)
{
    write_data(as_template);
    archive_.save(bytes);

    return true;
}

detail::stylesheet &excel_serializer::get_stylesheet()
{
    return workbook_.d_->stylesheet_;
}

} // namespace xlnt
