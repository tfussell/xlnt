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
#include <xlnt/packaging/document_properties.hpp>
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

bool load_workbook(xlnt::zip_file &archive, bool guess_types, bool data_only, xlnt::workbook &wb)
{
    wb.set_guess_types(guess_types);
    wb.set_data_only(data_only);

    if(!archive.has_file(xlnt::constants::ArcContentTypes()))
    {
        throw xlnt::invalid_file_exception("missing [Content Types].xml");
    }
    
    xlnt::manifest_serializer ms(wb.get_manifest());
	pugi::xml_document manifest_xml;
    manifest_xml.load(archive.read(xlnt::constants::ArcContentTypes()).c_str());
    ms.read_manifest(manifest_xml);

    if (ms.determine_document_type() != "excel")
    {
        throw xlnt::invalid_file_exception("package is not an OOXML SpreadsheetML");
    }

    wb.clear();
    
    if(archive.has_file(xlnt::constants::ArcCore()))
    {
        xlnt::workbook_serializer workbook_serializer_(wb);
		pugi::xml_document core_properties_xml;
		core_properties_xml.load(archive.read(xlnt::constants::ArcCore()).c_str());
        workbook_serializer_.read_properties_core(core_properties_xml);
    }
    
    if(archive.has_file(xlnt::constants::ArcApp()))
    {
        xlnt::workbook_serializer workbook_serializer_(wb);
		pugi::xml_document app_properties_xml;
		app_properties_xml.load(archive.read(xlnt::constants::ArcApp()).c_str());
        workbook_serializer_.read_properties_app(app_properties_xml);
    }

    xlnt::relationship_serializer relationship_serializer_(archive);
    auto root_relationships = relationship_serializer_.read_relationships("");
    
    for (const auto &relationship : root_relationships)
    {
        wb.create_root_relationship(relationship.get_id(), relationship.get_target_uri(), relationship.get_type());
    }
    
    auto workbook_relationships = relationship_serializer_.read_relationships(xlnt::constants::ArcWorkbook());

    for (const auto &relationship : workbook_relationships)
    {
        wb.create_relationship(relationship.get_id(), relationship.get_target_uri(), relationship.get_type());
    }

	pugi::xml_document xml;
    xml.load(archive.read(xlnt::constants::ArcWorkbook()).c_str());

    auto root_node = xml.child("workbook");

    auto workbook_pr_node = root_node.child("workbookPr");
    wb.get_properties().excel_base_date =
        (workbook_pr_node.attribute("date1904") && workbook_pr_node.attribute("date1904").value() != std::string("0"))
            ? xlnt::calendar::mac_1904
            : xlnt::calendar::windows_1900;

    if(archive.has_file(xlnt::constants::ArcSharedString()))
    {
        std::vector<xlnt::text> shared_strings;
		pugi::xml_document shared_strings_xml;
		shared_strings_xml.load(archive.read(xlnt::constants::ArcSharedString()).c_str());
		xlnt::shared_strings_serializer::read_shared_strings(shared_strings_xml, shared_strings);

        for (auto &shared_string : shared_strings)
        {
            wb.add_shared_string(shared_string, true);
        }
    }

    xlnt::detail::stylesheet stylesheet;
    xlnt::style_serializer style_serializer(stylesheet);
	pugi::xml_document style_xml;
	style_xml.load(archive.read(xlnt::constants::ArcStyles()).c_str());
    style_serializer.read_stylesheet(style_xml);

    auto sheets_node = root_node.child("sheets");

    for (auto sheet_node : sheets_node.children())
    {
        auto rel = wb.get_relationship(sheet_node.attribute("r:id").value());
        auto ws = wb.create_sheet(sheet_node.attribute("name").value(), rel);
        
        //TODO: this is really bad
        auto ws_filename = (rel.get_target_uri().substr(0, 3) != "xl/" ? "xl/" : "") + rel.get_target_uri();
        
        auto sheet_type = wb.get_manifest().get_override_type(ws_filename);
        
        if(rel.get_type() != xlnt::relationship::type::worksheet)
        {
            continue;
        }
        
        xlnt::worksheet_serializer worksheet_serializer(ws);
		pugi::xml_document worksheet_xml;
		worksheet_xml.load(archive.read(ws_filename).c_str());
        worksheet_serializer.read_worksheet(worksheet_xml, stylesheet);
    }

    if (archive.has_file("docProps/thumbnail.jpeg"))
    {
        auto thumbnail_data = archive.read("docProps/thumbnail.jpeg");
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

bool excel_serializer::load_workbook(const std::string &filename, bool guess_types, bool data_only)
{
    try
    {
        archive_.load(filename);
    }
    catch (std::runtime_error)
    {
        throw invalid_file_exception(filename);
    }

    return ::load_workbook(archive_, guess_types, data_only, workbook_);
}

bool excel_serializer::load_virtual_workbook(const std::vector<std::uint8_t> &bytes, bool guess_types, bool data_only)
{
    archive_.load(bytes);

    return ::load_workbook(archive_, guess_types, data_only, workbook_);
}

excel_serializer::excel_serializer(workbook &wb) : workbook_(wb)
{
}

void excel_serializer::write_data(bool /*as_template*/)
{
    relationship_serializer relationship_serializer_(archive_);
    relationship_serializer_.write_relationships(workbook_.get_root_relationships(), "");
    relationship_serializer_.write_relationships(workbook_.get_relationships(), constants::ArcWorkbook());

    pugi::xml_document properties_app_xml;
    workbook_serializer workbook_serializer_(workbook_);
    workbook_serializer_.write_properties_app(properties_app_xml);
    std::ostringstream ss;
    properties_app_xml.save(ss);
    archive_.writestr(constants::ArcApp(), ss.str());

    pugi::xml_document properties_core_xml;
    workbook_serializer_.write_properties_core(properties_core_xml);
    properties_core_xml.save(ss);
    archive_.writestr(constants::ArcCore(), ss.str());

    pugi::xml_document theme_xml;
    theme_serializer theme_serializer_;
    theme_serializer_.write_theme(workbook_.get_loaded_theme(), theme_xml);
    theme_xml.save(ss);
    archive_.writestr(constants::ArcTheme(), ss.str());

    if (!workbook_.get_shared_strings().empty())
    {
        const auto &strings = workbook_.get_shared_strings();
        pugi::xml_document shared_strings_xml;
        shared_strings_serializer strings_serializer;
        strings_serializer.write_shared_strings(strings, shared_strings_xml);
        shared_strings_xml.save(ss);

        archive_.writestr(constants::ArcSharedString(), ss.str());
    }

    pugi::xml_document workbook_xml;
    workbook_serializer_.write_workbook(workbook_xml);
    workbook_xml.save(ss);
    archive_.writestr(constants::ArcWorkbook(), ss.str());

    style_serializer style_serializer(workbook_.d_->stylesheet_);
    pugi::xml_document style_xml;
    style_serializer.write_stylesheet(style_xml);
    style_xml.save(ss);
    archive_.writestr(constants::ArcStyles(), ss.str());

    manifest_serializer manifest_serializer_(workbook_.get_manifest());
    pugi::xml_document manifest_xml;
    manifest_serializer_.write_manifest(manifest_xml);
    manifest_xml.save(ss);
    archive_.writestr(constants::ArcContentTypes(), ss.str());

    write_worksheets();
    
    if(!workbook_.get_thumbnail().empty())
    {
        const auto &thumbnail = workbook_.get_thumbnail();
        archive_.writestr("docProps/thumbnail.jpeg", std::string(thumbnail.begin(), thumbnail.end()));
    }
}

void excel_serializer::write_worksheets()
{
    std::size_t index = 0;

    for (auto ws : workbook_)
    {
        for (auto relationship : workbook_.get_relationships())
        {
            if (relationship.get_type() == relationship::type::worksheet &&
                workbook::index_from_ws_filename(relationship.get_target_uri()) == index)
            {
                worksheet_serializer serializer_(ws);
                std::string ws_filename = (relationship.get_target_uri().substr(0, 3) != "xl/" ? "xl/" : "") + relationship.get_target_uri();
                std::ostringstream ss;
                pugi::xml_document worksheet_xml;
                serializer_.write_worksheet(worksheet_xml);
                worksheet_xml.save(ss);
                archive_.writestr(ws_filename, ss.str());

                break;
            }
        }

        index++;
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

bool excel_serializer::save_workbook(const std::string &filename, bool as_template)
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
