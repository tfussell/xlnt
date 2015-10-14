#include <sstream>

#include <xlnt/common/exceptions.hpp>
#include <xlnt/common/relationship.hpp>
#include <xlnt/common/zip_file.hpp>
#include <xlnt/workbook/document_properties.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/worksheet/range_reference.hpp>
#include <xlnt/worksheet/worksheet.hpp>
#include <xlnt/writer/manifest_writer.hpp>
#include <xlnt/writer/relationship_writer.hpp>
#include <xlnt/writer/theme_writer.hpp>
#include <xlnt/writer/workbook_writer.hpp>

#include "constants.hpp"
#include "detail/include_pugixml.hpp"

namespace {

std::string to_xml(xlnt::document_properties &props)
{
    return "";
}

} // namespace

namespace xlnt {

excel_writer::excel_writer(workbook &wb) : wb_(wb), style_writer_(wb_)
{
}

void excel_writer::save(const std::string &filename, bool as_template)
{
    zip_file archive;
    write_data(archive, as_template);
    archive.save(filename);
}

void excel_writer::write_data(zip_file &archive, bool as_template)
{
    archive.writestr(constants::ArcRootRels, write_root_rels(wb_));
    archive.writestr(constants::ArcWorkbookRels, write_workbook_rels(wb_));
    archive.writestr(constants::ArcApp, write_properties_app(wb_));
    archive.writestr(constants::ArcCore, to_xml(wb_.get_properties()));
    
    if(wb_.has_loaded_theme())
    {
        archive.writestr(constants::ArcTheme, wb_.get_loaded_theme());
    }
    else
    {
        archive.writestr(constants::ArcTheme, write_theme());
    }
    
    archive.writestr(constants::ArcWorkbook, write_workbook(wb_));

    /*
    if(wb_.has_vba_archive())
    {
        auto &vba_archive = wb_.get_vba_archive();
        
        for(auto name : vba_archive.namelist())
        {
            for(auto s : constants::ArcVba)
            {
                if(match(s, name))
                {
                    archive.writestr(name, vba_archive.read(name));
                    break;
                }
            }
        }
    }
     */

    write_charts(archive);
    write_images(archive);
    write_worksheets(archive);
    write_chartsheets(archive);
    write_string_table(archive);
    write_external_links(archive);
    archive.writestr(constants::ArcStyles, style_writer_.write_table());
    auto manifest = write_content_types(wb_, as_template);
    archive.writestr(constants::ArcContentTypes, manifest);
}

void excel_writer::write_string_table(zip_file &archive)
{
    
}

void excel_writer::write_images(zip_file &archive)
{
    
}

void excel_writer::write_charts(zip_file &archive)
{
    
}

void excel_writer::write_chartsheets(zip_file &archive)
{
    
}

void excel_writer::write_worksheets(zip_file &archive)
{
    
}

void excel_writer::write_external_links(zip_file &archive)
{
    
}

std::string write_properties_app(const workbook &wb)
{
    pugi::xml_document doc;
    auto root_node = doc.append_child("Properties");
    root_node.append_attribute("xmlns").set_value("http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");
    root_node.append_attribute("xmlns:vt").set_value("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
    
    root_node.append_child("Application").text().set("Microsoft Excel");
    root_node.append_child("DocSecurity").text().set("0");
    root_node.append_child("ScaleCrop").text().set("false");
    root_node.append_child("Company");
    root_node.append_child("LinksUpToDate").text().set("false");
    root_node.append_child("SharedDoc").text().set("false");
    root_node.append_child("HyperlinksChanged").text().set("false");
    root_node.append_child("AppVersion").text().set("12.0000");
    
    auto heading_pairs_node = root_node.append_child("HeadingPairs");
    auto heading_pairs_vector_node = heading_pairs_node.append_child("vt:vector");
    heading_pairs_vector_node.append_attribute("baseType").set_value("variant");
    heading_pairs_vector_node.append_attribute("size").set_value("2");
    heading_pairs_vector_node.append_child("vt:variant").append_child("vt:lpstr").text().set("Worksheets");
    heading_pairs_vector_node.append_child("vt:variant").append_child("vt:i4").text().set(std::to_string(wb.get_sheet_names().size()).c_str());
    
    auto titles_of_parts_node = root_node.append_child("TitlesOfParts");
    auto titles_of_parts_vector_node = titles_of_parts_node.append_child("vt:vector");
    titles_of_parts_vector_node.append_attribute("baseType").set_value("lpstr");
    titles_of_parts_vector_node.append_attribute("size").set_value(std::to_string(wb.get_sheet_names().size()).c_str());
    
    for(auto ws : wb)
    {
        titles_of_parts_vector_node.append_child("vt:lpstr").text().set(ws.get_title().c_str());
    }
    
    std::stringstream ss;
    doc.save(ss);
    
    return ss.str();
}

std::string write_root_rels(const workbook &)
{
    std::vector<relationship> relationships;
    
    relationships.push_back(relationship(relationship::type::extended_properties, "rId3", "docProps/app.xml"));
    relationships.push_back(relationship(relationship::type::core_properties, "rId2", "docProps/core.xml"));
    relationships.push_back(relationship(relationship::type::office_document, "rId1", "xl/workbook.xml"));
    
    return write_relationships(relationships, "");
}

std::string write_workbook(const workbook &wb)
{
    std::size_t num_visible = 0;
    
    for(auto ws : wb)
    {
        if(ws.get_page_setup().get_sheet_state() == xlnt::page_setup::sheet_state::visible)
        {
            num_visible++;
        }
    }
    
    if(num_visible == 0)
    {
        throw value_error();
    }
    
    pugi::xml_document doc;
    auto root_node = doc.append_child("workbook");
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
    
    for(auto relationship : wb.get_relationships())
    {
        if(relationship.get_type() == relationship::type::worksheet)
        {
            std::string sheet_index_string = relationship.get_target_uri();
            sheet_index_string = sheet_index_string.substr(0, sheet_index_string.find('.'));
            sheet_index_string = sheet_index_string.substr(sheet_index_string.find_last_of('/'));
            auto iter = sheet_index_string.end();
            iter--;
            while (isdigit(*iter)) iter--;
            auto first_digit = iter - sheet_index_string.begin();
            sheet_index_string = sheet_index_string.substr(first_digit + 1);
            std::size_t sheet_index = std::stoi(sheet_index_string) - 1;
            
            auto ws = wb.get_sheet_by_index(sheet_index);
            
            auto sheet_node = sheets_node.append_child("sheet");
            sheet_node.append_attribute("name").set_value(ws.get_title().c_str());
            sheet_node.append_attribute("r:id").set_value(relationship.get_id().c_str());
            sheet_node.append_attribute("sheetId").set_value(std::to_string(sheet_index + 1).c_str());
            
            if(ws.has_auto_filter())
            {
                auto defined_name_node = defined_names_node.append_child("definedName");
                defined_name_node.append_attribute("name").set_value("_xlnm._FilterDatabase");
                defined_name_node.append_attribute("hidden").set_value(1);
                defined_name_node.append_attribute("localSheetId").set_value(0);
                std::string name = "'" + ws.get_title() + "'!" + range_reference::make_absolute(ws.get_auto_filter()).to_string();
                defined_name_node.text().set(name.c_str());
            }
        }
    }
    
    auto calc_pr_node = root_node.append_child("calcPr");
    calc_pr_node.append_attribute("calcId").set_value("124519");
    calc_pr_node.append_attribute("calcMode").set_value("auto");
    calc_pr_node.append_attribute("fullCalcOnLoad").set_value("1");
    
    std::stringstream ss;
    doc.save(ss);
    
    return ss.str();
}

std::string write_workbook_rels(const workbook &wb)
{
    return write_relationships(wb.get_relationships(), "xl/");
}

std::string write_theme()
{
    return "";
}

bool save_workbook(workbook &wb, const std::string &filename, bool as_template)
{
    excel_writer writer(wb);
    writer.save(filename, as_template);
    return true;
}

std::vector<std::uint8_t> save_virtual_workbook(xlnt::workbook &wb, bool as_template)
{
    excel_writer writer(wb);
    std::vector<std::uint8_t> buffer;
    zip_file archive(buffer);
    writer.write_data(archive, as_template);
    return buffer;
}

} // namespace xlnt
