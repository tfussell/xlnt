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
#include <pugixml.hpp>

#include <detail/constants.hpp>
#include <detail/workbook_serializer.hpp>
#include <xlnt/packaging/app_properties.hpp>
#include <xlnt/packaging/document_properties.hpp>
#include <xlnt/packaging/manifest.hpp>
#include <xlnt/packaging/relationship.hpp>
#include <xlnt/utils/datetime.hpp>
#include <xlnt/utils/exceptions.hpp>
#include <xlnt/workbook/named_range.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/workbook/worksheet_iterator.hpp>
#include <xlnt/worksheet/range_reference.hpp>
#include <xlnt/worksheet/worksheet.hpp>

namespace {

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
    return std::to_string(dt.year) + "-" + fill(std::to_string(dt.month)) + "-" + fill(std::to_string(dt.day)) + "T" +
           fill(std::to_string(dt.hour)) + ":" + fill(std::to_string(dt.minute)) + ":" +
           fill(std::to_string(dt.second)) + "Z";
}

} // namespace

namespace xlnt {

workbook_serializer::workbook_serializer(workbook &wb) : workbook_(wb)
{
}

void workbook_serializer::read_properties_core(const pugi::xml_document &xml)
{
    auto &props = workbook_.get_properties();
    auto root_node = xml.child("cp:coreProperties");

    props.excel_base_date = calendar::windows_1900;

    if (root_node.child("dc:creator"))
    {
        props.creator = root_node.child("dc:creator").text().get();
    }
    
    if (root_node.child("cp:lastModifiedBy"))
    {
        props.last_modified_by = root_node.child("cp:lastModifiedBy").text().get();
    }
    
    if (root_node.child("dcterms:created"))
    {
        std::string created_string = root_node.child("dcterms:created").text().get();
        props.created = w3cdtf_to_datetime(created_string);
    }
    
    if (root_node.child("dcterms:modified"))
    {
        std::string modified_string = root_node.child("dcterms:modified").text().get();
        props.modified = w3cdtf_to_datetime(modified_string);
    }
}

void workbook_serializer::read_properties_app(const pugi::xml_document &xml)
{
    auto &props = workbook_.get_app_properties();
    auto root_node = xml.child("Properties");

    if(root_node.child("Application"))
    {
        props.application = root_node.child("Application").text().get();
    }

    if(root_node.child("DocSecurity"))
    {
        props.doc_security = std::stoi(root_node.child("DocSecurity").text().get());
    }

    if(root_node.child("ScaleCrop"))
    {
        props.scale_crop = root_node.child("ScaleCrop").text().get() == std::string("true");
    }

    if(root_node.child("Company"))
    {
        props.company = root_node.child("Company").text().get();
    }

    if(root_node.child("ScaleCrop"))
    {
        props.links_up_to_date = root_node.child("ScaleCrop").text().get() == std::string("true");
    }

    if(root_node.child("SharedDoc"))
    {
        props.shared_doc = root_node.child("SharedDoc").text().get() == std::string("true");
    }

    if(root_node.child("HyperlinksChanged"))
    {
        props.hyperlinks_changed = root_node.child("HyperlinksChanged").text().get() == std::string("true");
    }

    if(root_node.child("AppVersion"))
    {
        props.app_version = root_node.child("AppVersion").text().get();
    }
}

void workbook_serializer::write_properties_core(pugi::xml_document &xml) const
{
    auto &props = workbook_.get_properties();
    auto root_node = xml.append_child("cp:coreProperties");

    root_node.append_attribute("xmlns:cp").set_value("http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
    root_node.append_attribute("xmlns:dc").set_value("http://purl.org/dc/elements/1.1/");
    root_node.append_attribute("xmlns:dcmitype").set_value("http://purl.org/dc/dcmitype/");
    root_node.append_attribute("xmlns:dcterms").set_value("http://purl.org/dc/terms/");
    root_node.append_attribute("xmlns:xsi").set_value("http://www.w3.org/2001/XMLSchema-instance");

    root_node.append_child("dc:creator").text().set(props.creator.c_str());
    root_node.append_child("cp:lastModifiedBy").text().set(props.last_modified_by.c_str());
    root_node.append_child("dcterms:created").text().set(datetime_to_w3cdtf(props.created).c_str());
    root_node.child("dcterms:created").append_attribute("xsi:type").set_value("dcterms:W3CDTF");
    root_node.append_child("dcterms:modified").text().set(datetime_to_w3cdtf(props.modified).c_str());
    root_node.child("dcterms:modified").append_attribute("xsi:type").set_value("dcterms:W3CDTF");
    root_node.append_child("dc:title").text().set(props.title.c_str());
    root_node.append_child("dc:description");
    root_node.append_child("dc:subject");
    root_node.append_child("cp:keywords");
    root_node.append_child("cp:category");
}

void workbook_serializer::write_properties_app(pugi::xml_document &xml) const
{
    auto root_node = xml.append_child("Properties");

    root_node.append_attribute("xmlns").set_value("http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");
    root_node.append_attribute("xmlns:vt").set_value("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
    
    auto &properties = workbook_.get_app_properties();

    root_node.append_child("Application").text().set(properties.application.c_str());
    root_node.append_child("DocSecurity").text().set(std::to_string(properties.doc_security).c_str());
    root_node.append_child("ScaleCrop").text().set(properties.scale_crop ? "true" : "false");
    
    auto company_node = root_node.append_child("Company");
    
    if (!properties.company.empty())
    {
        company_node.text().set(properties.company.c_str());
    }
    
    root_node.append_child("LinksUpToDate").text().set(properties.links_up_to_date ? "true" : "false");
    root_node.append_child("SharedDoc").text().set(properties.shared_doc ? "true" : "false");
    root_node.append_child("HyperlinksChanged").text().set(properties.hyperlinks_changed ? "true" : "false");
    root_node.append_child("AppVersion").text().set(properties.app_version.c_str());

    // TODO what is this stuff?
    
    auto heading_pairs_node = root_node.append_child("HeadingPairs");
    auto heading_pairs_vector_node = heading_pairs_node.append_child("vt:vector");
    heading_pairs_vector_node.append_attribute("size").set_value("2");
    heading_pairs_vector_node.append_attribute("baseType").set_value("variant");
    heading_pairs_vector_node.append_child("vt:variant").append_child("vt:lpstr").text().set("Worksheets");
    heading_pairs_vector_node.append_child("vt:variant")
        .append_child("vt:i4")
        .text().set(std::to_string(workbook_.get_sheet_names().size()).c_str());

    auto titles_of_parts_node = root_node.append_child("TitlesOfParts");
    auto titles_of_parts_vector_node = titles_of_parts_node.append_child("vt:vector");
    titles_of_parts_vector_node.append_attribute("size").set_value(std::to_string(workbook_.get_sheet_names().size()).c_str());
    titles_of_parts_vector_node.append_attribute("baseType").set_value("lpstr");

    for (auto ws : workbook_)
    {
        titles_of_parts_vector_node.append_child("vt:lpstr").text().set(ws.get_title().c_str());
    }
}

void workbook_serializer::write_workbook(pugi::xml_document &xml) const
{
    std::size_t num_visible = 0;

    for (auto ws : workbook_)
    {
        if (ws.get_page_setup().get_sheet_state() == sheet_state::visible)
        {
            num_visible++;
        }
    }

    if (num_visible == 0)
    {
        throw xlnt::value_error();
    }

    auto root_node = xml.append_child("workbook");

    root_node.append_attribute("xmlns").set_value("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    root_node.append_attribute("xmlns:r").set_value("http://schemas.openxmlformats.org/officeDocument/2006/relationships");

    auto file_version_node = root_node.append_child("fileVersion");
    file_version_node.append_attribute("appName").set_value("xl");
    file_version_node.append_attribute("lastEdited").set_value("4");
    file_version_node.append_attribute("lowestEdited").set_value("4");
    file_version_node.append_attribute("rupBuild").set_value("4505");

    auto workbook_pr_node = root_node.append_child("workbookPr");
    workbook_pr_node.append_attribute("codeName").set_value("ThisWorkbook");
    workbook_pr_node.append_attribute("defaultThemeVersion").set_value("124226");
    workbook_pr_node.append_attribute("date1904").set_value(workbook_.get_properties().excel_base_date == calendar::mac_1904 ? "1" : "0");

    auto book_views_node = root_node.append_child("bookViews");
    auto workbook_view_node = book_views_node.append_child("workbookView");
    workbook_view_node.append_attribute("activeTab").set_value("0");
    workbook_view_node.append_attribute("autoFilterDateGrouping").set_value("1");
    workbook_view_node.append_attribute("firstSheet").set_value("0");
    workbook_view_node.append_attribute("minimized").set_value("0");
    workbook_view_node.append_attribute("showHorizontalScroll").set_value("1");
    workbook_view_node.append_attribute("showSheetTabs").set_value("1");
    workbook_view_node.append_attribute("showVerticalScroll").set_value("1");
    workbook_view_node.append_attribute("tabRatio").set_value("600");
    workbook_view_node.append_attribute("visibility").set_value("visible");

    auto sheets_node = root_node.append_child("sheets");
    auto defined_names_node = root_node.append_child("definedNames");

    for (const auto &relationship : workbook_.get_relationships())
    {
        if (relationship.get_type() == relationship::type::worksheet)
        {
            // TODO: this is ugly
            std::string sheet_index_string = relationship.get_target_uri();
            sheet_index_string = sheet_index_string.substr(0, sheet_index_string.find('.'));
            sheet_index_string = sheet_index_string.substr(sheet_index_string.find_last_of('/'));
            auto iter = sheet_index_string.end();
            iter--;
            while (isdigit(*iter))
                iter--;
            auto first_digit = iter - sheet_index_string.begin();
            sheet_index_string = sheet_index_string.substr(static_cast<std::string::size_type>(first_digit + 1));
            std::size_t sheet_index = static_cast<std::size_t>(std::stoll(sheet_index_string) - 1);

            auto ws = workbook_.get_sheet_by_index(sheet_index);

            auto sheet_node = sheets_node.append_child("sheet");
            sheet_node.append_attribute("name").set_value(ws.get_title().c_str());
            sheet_node.append_attribute("sheetId").set_value(std::to_string(sheet_index + 1).c_str());
            sheet_node.append_attribute("r:id").set_value(relationship.get_id().c_str());

            if (ws.has_auto_filter())
            {
                auto defined_name_node = defined_names_node.append_child("definedName");
                defined_name_node.append_attribute("name").set_value("_xlnm._FilterDatabase");
                defined_name_node.append_attribute("hidden").set_value("1");
                defined_name_node.append_attribute("localSheetId").set_value("0");
                std::string name =
                    "'" + ws.get_title() + "'!" + range_reference::make_absolute(ws.get_auto_filter()).to_string();
                defined_name_node.text().set(name.c_str());
            }
        }
    }

    auto calc_pr_node = root_node.append_child("calcPr");
    calc_pr_node.append_attribute("calcId").set_value("124519");
    calc_pr_node.append_attribute("calcMode").set_value("auto");
    calc_pr_node.append_attribute("fullCalcOnLoad").set_value("1");
}

void workbook_serializer::write_named_ranges(pugi::xml_node node) const
{
    for (auto &named_range : workbook_.get_named_ranges())
    {
        node.append_child(named_range.get_name().c_str());
    }
}

} // namespace xlnt
