#include <algorithm>
#include <array>
#include <cmath>
#include <sstream>
#include <string>
#include <unordered_map>

#include <pugixml.hpp>

#include <xlnt/writer/writer.hpp>
#include <xlnt/cell/cell.hpp>
#include <xlnt/cell/value.hpp>
#include <xlnt/worksheet/range.hpp>
#include <xlnt/worksheet/range_reference.hpp>
#include <xlnt/worksheet/worksheet.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/common/relationship.hpp>
#include <xlnt/workbook/document_properties.hpp>

#include "constants.hpp"

namespace xlnt {

std::string writer::write_shared_strings(const std::vector<std::string> &string_table)
{
    pugi::xml_document doc;
    auto root_node = doc.append_child("sst");
    root_node.append_attribute("xmlns").set_value("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    root_node.append_attribute("uniqueCount").set_value((int)string_table.size());
    
    for(auto string : string_table)
    {
        root_node.append_child("si").append_child("t").text().set(string.c_str());
    }
    
    std::stringstream ss;
    doc.save(ss);
    
    return ss.str();
}

std::string fill(const std::string &string, std::size_t length = 2)
{
    if(string.size() >= length)
    {
        return string;
    }

    return std::string(length - string.size(), '0') + string;
}

std::string datetime_to_w3cdtf(const datetime &dt)
{
    return std::to_string(dt.year) + "-" + fill(std::to_string(dt.month)) + "-" + fill(std::to_string(dt.day)) + "T" + fill(std::to_string(dt.hour)) + ":" + fill(std::to_string(dt.minute)) + ":" + fill(std::to_string(dt.second)) + "Z";
}

std::string writer::write_properties_core(const document_properties &prop)
{
    pugi::xml_document doc;
    auto root_node = doc.append_child("cp:coreProperties");
    root_node.append_attribute("xmlns:cp").set_value("http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
    root_node.append_attribute("xmlns:dc").set_value("http://purl.org/dc/elements/1.1/");
    root_node.append_attribute("xmlns:dcmitype").set_value("http://purl.org/dc/dcmitype/");
    root_node.append_attribute("xmlns:dcterms").set_value("http://purl.org/dc/terms/");
    root_node.append_attribute("xmlns:xsi").set_value("http://www.w3.org/2001/XMLSchema-instance");

    root_node.append_child("dc:creator").text().set(prop.creator.c_str());
    root_node.append_child("cp:lastModifiedBy").text().set(prop.last_modified_by.c_str());
    root_node.append_child("dcterms:created").text().set(datetime_to_w3cdtf(prop.created).c_str());
    root_node.child("dcterms:created").append_attribute("xsi:type").set_value("dcterms:W3CDTF");
    root_node.append_child("dcterms:modified").text().set(datetime_to_w3cdtf(prop.modified).c_str());
    root_node.child("dcterms:modified").append_attribute("xsi:type").set_value("dcterms:W3CDTF");
    root_node.append_child("dc:title").text().set(prop.title.c_str());
    root_node.append_child("dc:description");
    root_node.append_child("dc:subject");
    root_node.append_child("cp:keywords");
    root_node.append_child("cp:category");

    std::stringstream ss;
    doc.save(ss);

    return ss.str();
}

std::string writer::write_properties_app(const workbook &wb)
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

std::string writer::write_workbook_rels(const workbook &wb)
{
    return write_relationships(wb.get_relationships());
}

std::string writer::write_worksheet_rels(worksheet ws)
{
	return write_relationships(ws.get_relationships());
}


std::string writer::write_workbook(const workbook &wb)
{
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
            std::string sheet_index_string = relationship.get_target_uri().substr(16);
            std::size_t sheet_index = std::stoi(sheet_index_string.substr(0, sheet_index_string.find('.'))) - 1;
            
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

std::string writer::write_worksheet(worksheet ws, const std::vector<std::string> &string_table, const std::unordered_map<std::size_t, std::string> &style_id_by_hash)
{
    ws.get_cell("A1");

    pugi::xml_document doc;
    auto root_node = doc.append_child("worksheet");
    root_node.append_attribute("xmlns").set_value(constants::Namespaces.at("spreadsheetml").c_str());
    root_node.append_attribute("xmlns:r").set_value(constants::Namespaces.at("r").c_str());
    auto sheet_pr_node = root_node.append_child("sheetPr");
    auto outline_pr_node = sheet_pr_node.append_child("outlinePr");
    if(!ws.get_page_setup().is_default())
    {
        auto page_set_up_pr_node = sheet_pr_node.append_child("pageSetUpPr");
        page_set_up_pr_node.append_attribute("fitToPage").set_value(ws.get_page_setup().fit_to_page() ? 1 : 0);
    }
    outline_pr_node.append_attribute("summaryBelow").set_value(1);
    outline_pr_node.append_attribute("summaryRight").set_value(1);
    auto dimension_node = root_node.append_child("dimension");
    dimension_node.append_attribute("ref").set_value(ws.calculate_dimension().to_string().c_str());
    auto sheet_views_node = root_node.append_child("sheetViews");
    auto sheet_view_node = sheet_views_node.append_child("sheetView");
    sheet_view_node.append_attribute("workbookViewId").set_value(0);
    
    std::string active_pane = "bottomRight";

    if(ws.has_frozen_panes())
    {
        auto pane_node = sheet_view_node.append_child("pane");

        if(ws.get_frozen_panes().get_column_index() > 0)
        {
            pane_node.append_attribute("xSplit").set_value(ws.get_frozen_panes().get_column_index());
            active_pane = "topRight";
        }

        if(ws.get_frozen_panes().get_row_index() > 0)
        {
            pane_node.append_attribute("ySplit").set_value(ws.get_frozen_panes().get_row_index());
            active_pane = "bottomLeft";
        }

        if(ws.get_frozen_panes().get_row_index() > 0 && ws.get_frozen_panes().get_column_index() > 0)
        {
            auto top_right_node = sheet_view_node.append_child("selection");
            top_right_node.append_attribute("pane").set_value("topRight");
            auto bottom_left_node = sheet_view_node.append_child("selection");
            bottom_left_node.append_attribute("pane").set_value("bottomLeft");
            active_pane = "bottomRight";
        }

        pane_node.append_attribute("topLeftCell").set_value(ws.get_frozen_panes().to_string().c_str());
        pane_node.append_attribute("activePane").set_value(active_pane.c_str());
        pane_node.append_attribute("state").set_value("frozen");
    }

    auto selection_node = sheet_view_node.append_child("selection");
    if(ws.has_frozen_panes())
    {
        if(ws.get_frozen_panes().get_row_index() > 0 && ws.get_frozen_panes().get_column_index() > 0)
        {
            selection_node.append_attribute("pane").set_value("bottomRight");
        }
        else if(ws.get_frozen_panes().get_row_index() > 0)
        {
            selection_node.append_attribute("pane").set_value("bottomLeft");
        }
        else if(ws.get_frozen_panes().get_column_index() > 0)
        {
            selection_node.append_attribute("pane").set_value("topRight");
        }
    }
    std::string active_cell = "A1";
    selection_node.append_attribute("activeCell").set_value(active_cell.c_str());
    selection_node.append_attribute("sqref").set_value(active_cell.c_str());
    
    auto sheet_format_pr_node = root_node.append_child("sheetFormatPr");
    sheet_format_pr_node.append_attribute("baseColWidth").set_value(10);
    sheet_format_pr_node.append_attribute("defaultRowHeight").set_value(15);
    
    std::vector<int> styled_columns;
    
    if(!style_id_by_hash.empty())
    {
        for(auto row : ws.rows())
        {
            for(auto cell : row)
            {
                if(cell.has_style())
                {
                    styled_columns.push_back(xlnt::cell_reference::column_index_from_string(cell.get_column()));
                }
            }
        }
        
        auto cols_node = root_node.append_child("cols");
        std::sort(styled_columns.begin(), styled_columns.end());
        for(auto column : styled_columns)
        {
            auto col_node = cols_node.append_child("col");
            col_node.append_attribute("min").set_value(column);
            col_node.append_attribute("max").set_value(column);
            col_node.append_attribute("style").set_value(1);
        }
    }

    std::unordered_map<std::string, std::string> hyperlink_references;
    
    auto sheet_data_node = root_node.append_child("sheetData");
    for(auto row : ws.rows())
    {
        row_t min = (int)row.num_cells();
        row_t max = 0;
        bool any_non_null = false;
        
        for(auto cell : row)
        {
            min = std::min(min, cell_reference::column_index_from_string(cell.get_column()));
            max = std::max(max, cell_reference::column_index_from_string(cell.get_column()));
            
            if(!cell.garbage_collectible())
            {
                any_non_null = true;
            }
        }
        
        if(!any_non_null)
        {
            continue;
        }
        
        auto row_node = sheet_data_node.append_child("row");
        row_node.append_attribute("r").set_value(row.front().get_row());
        
        row_node.append_attribute("spans").set_value((std::to_string(min) + ":" + std::to_string(max)).c_str());
        if(ws.has_row_properties(row.front().get_row()))
        {
            row_node.append_attribute("customHeight").set_value(1);
            auto height = ws.get_row_properties(row.front().get_row()).height;
            if(height == std::floor(height))
            {
                row_node.append_attribute("ht").set_value((std::to_string((int)height) + ".0").c_str());
            }
            else
            {
                row_node.append_attribute("ht").set_value(height);
            }
        }
        //row_node.append_attribute("x14ac:dyDescent").set_value(0.25);
        
        for(auto cell : row)
        {
            if(!cell.garbage_collectible())
            {
                if(cell.has_hyperlink())
                {
                    hyperlink_references[cell.get_hyperlink().get_id()] = cell.get_reference().to_string();
                }

                auto cell_node = row_node.append_child("c");
                cell_node.append_attribute("r").set_value(cell.get_reference().to_string().c_str());
                
                if(cell.get_value().is(value::type::string))
                {
                    if(cell.has_formula())
                    {
                        cell_node.append_attribute("t").set_value("str");
                        cell_node.append_child("f").text().set(cell.get_formula().c_str());
                        cell_node.append_child("v").text().set(cell.get_value().to_string().c_str());
                        continue;
                    }

                    int match_index = -1;
                    for(int i = 0; i < (int)string_table.size(); i++)
                    {
                        if(string_table[i] == cell.get_value().as<std::string>())
                        {
                            match_index = i;
                            break;
                        }
                    }
                    
                    if(match_index == -1)
                    {
                        if(cell.get_value().as<std::string>().empty())
                        {
                            cell_node.append_attribute("t").set_value("s");
                        }
                        else
                        {
                            cell_node.append_attribute("t").set_value("inlineStr");
                            auto inline_string_node = cell_node.append_child("is");
                            inline_string_node.append_child("t").text().set(cell.get_value().as<std::string>().c_str());
                        }
                    }
                    else
                    {
                        cell_node.append_attribute("t").set_value("s");
                        auto value_node = cell_node.append_child("v");
                        value_node.text().set(match_index);
                    }
                }
                else
                {
                    if(!cell.get_value().is(value::type::null))
                    {
                        if(cell.get_value().is(value::type::boolean))
                        {
                            cell_node.append_attribute("t").set_value("b");
                            auto value_node = cell_node.append_child("v");
                            value_node.text().set(cell.get_value().as<bool>() ? 1 : 0);
                        }
                        else if(cell.get_value().is(value::type::numeric))
                        {
                            if(cell.has_formula())
                            {
                                cell_node.append_child("f").text().set(cell.get_formula().c_str());
                                cell_node.append_child("v").text().set(cell.get_value().to_string().c_str());
                                continue;
                            }

                            cell_node.append_attribute("t").set_value("n");
                            auto value_node = cell_node.append_child("v");
                            if(cell.get_value().is_integral())
                            {
                                value_node.text().set(cell.get_value().as<long long>());
                            }
                            else
                            {
                                value_node.text().set(cell.get_value().as<double>());
                            }
                        }
                    }
                    else if(cell.has_formula())
                    {
                        cell_node.append_child("f").text().set(cell.get_formula().c_str());
                        cell_node.append_child("v");
                        continue;
                    }
                }
                
                if(cell.has_style())
                {
                    cell_node.append_attribute("s").set_value(1);
                }
            }
        }
    }

    if(ws.has_auto_filter())
    {
        auto auto_filter_node = root_node.append_child("autoFilter");
        auto_filter_node.append_attribute("ref").set_value(ws.get_auto_filter().to_string().c_str());
    }
    
    if(!ws.get_merged_ranges().empty())
    {
        auto merge_cells_node = root_node.append_child("mergeCells");
        merge_cells_node.append_attribute("count").set_value((unsigned int)ws.get_merged_ranges().size());
        
        for(auto merged_range : ws.get_merged_ranges())
        {
            auto merge_cell_node = merge_cells_node.append_child("mergeCell");
            merge_cell_node.append_attribute("ref").set_value(merged_range.to_string().c_str());
        }
    }

    if(!ws.get_relationships().empty())
    {
        auto hyperlinks_node = root_node.append_child("hyperlinks");

        for(auto relationship : ws.get_relationships())
        {
            auto hyperlink_node = hyperlinks_node.append_child("hyperlink");
            hyperlink_node.append_attribute("display").set_value(relationship.get_target_uri().c_str());
            hyperlink_node.append_attribute("ref").set_value(hyperlink_references.at(relationship.get_id()).c_str());
            hyperlink_node.append_attribute("r:id").set_value(relationship.get_id().c_str());
        }
    }
    
    if(!ws.get_page_setup().is_default())
    {
        auto print_options_node = root_node.append_child("printOptions");
        print_options_node.append_attribute("horizontalCentered").set_value(ws.get_page_setup().get_horizontal_centered() ? 1 : 0);
        print_options_node.append_attribute("verticalCentered").set_value(ws.get_page_setup().get_vertical_centered() ? 1 : 0);
    }

    auto page_margins_node = root_node.append_child("pageMargins");
        
    page_margins_node.append_attribute("left").set_value(ws.get_page_margins().get_left());
    page_margins_node.append_attribute("right").set_value(ws.get_page_margins().get_right());
    page_margins_node.append_attribute("top").set_value(ws.get_page_margins().get_top());
    page_margins_node.append_attribute("bottom").set_value(ws.get_page_margins().get_bottom());
    page_margins_node.append_attribute("header").set_value(ws.get_page_margins().get_header());
    page_margins_node.append_attribute("footer").set_value(ws.get_page_margins().get_footer());
    
    if(!ws.get_page_setup().is_default())
    {
        auto page_setup_node = root_node.append_child("pageSetup");
        
        std::string orientation_string = ws.get_page_setup().get_orientation() == page_setup::orientation::landscape ? "landscape" : "portrait";
        page_setup_node.append_attribute("orientation").set_value(orientation_string.c_str());
        page_setup_node.append_attribute("paperSize").set_value((int)ws.get_page_setup().get_paper_size());
        page_setup_node.append_attribute("fitToHeight").set_value(ws.get_page_setup().fit_to_height() ? 1 : 0);
        page_setup_node.append_attribute("fitToWidth").set_value(ws.get_page_setup().fit_to_width() ? 1 : 0);
    }
    
    if(!ws.get_header_footer().is_default())
    {
        auto header_footer_node = root_node.append_child("headerFooter");
        auto odd_header_node = header_footer_node.append_child("oddHeader");
        std::string header_text = "&L&\"Calibri,Regular\"&K000000Left Header Text&C&\"Arial,Regular\"&6&K445566Center Header Text&R&\"Arial,Bold\"&8&K112233Right Header Text";
        odd_header_node.text().set(header_text.c_str());
        auto odd_footer_node = header_footer_node.append_child("oddFooter");
        std::string footer_text = "&L&\"Times New Roman,Regular\"&10&K445566Left Footer Text_x000D_And &D and &T&C&\"Times New Roman,Bold\"&12&K778899Center Footer Text &Z&F on &A&R&\"Times New Roman,Italic\"&14&KAABBCCRight Footer Text &P of &N";
        odd_footer_node.text().set(footer_text.c_str());
    }
    
    std::stringstream ss;
    doc.save(ss);
    
    return ss.str();
}
    
std::string writer::write_content_types(const workbook &wb)
{
    pugi::xml_document doc;
    auto root_node = doc.append_child("Types");
    root_node.append_attribute("xmlns").set_value(constants::Namespaces.at("content-types").c_str());
    
    for(auto type : wb.get_content_types())
    {
		pugi::xml_node type_node;

		if (type.is_default)
		{
			type_node = root_node.append_child("Default");
			type_node.append_attribute("Extension").set_value(type.extension.c_str());
		}
		else
		{
			type_node = root_node.append_child("Override");
			type_node.append_attribute("PartName").set_value(type.part_name.c_str());
		}

		type_node.append_attribute("ContentType").set_value(type.type.c_str());
    }
    
    std::stringstream ss;
    doc.save(ss);
    return ss.str();
}

std::string xlnt::writer::write_root_rels()
{
	std::vector<relationship> relationships;

	relationships.push_back(relationship(relationship::type::extended_properties, "rId3", "docProps/app.xml"));
	relationships.push_back(relationship(relationship::type::core_properties, "rId2", "docProps/core.xml"));
	relationships.push_back(relationship(relationship::type::office_document, "rId1", "xl/workbook.xml"));

	return write_relationships(relationships);
}

std::string writer::write_relationships(const std::vector<relationship> &relationships)
{
    pugi::xml_document doc;
    auto root_node = doc.append_child("Relationships");
    root_node.append_attribute("xmlns").set_value(constants::Namespaces.at("relationships").c_str());
    
    for(auto relationship : relationships)
    {
        auto app_props_node = root_node.append_child("Relationship");
        app_props_node.append_attribute("Id").set_value(relationship.get_id().c_str());
        app_props_node.append_attribute("Target").set_value(relationship.get_target_uri().c_str());
        app_props_node.append_attribute("Type").set_value(relationship.get_type_string().c_str());
        if(relationship.get_target_mode() == target_mode::external)
        {
            app_props_node.append_attribute("TargetMode").set_value("External");
        }
    }
    
    std::stringstream ss;
    doc.save(ss);
    return ss.str();
}
    
std::string writer::write_theme()
{
     pugi::xml_document doc;
     auto theme_node = doc.append_child("a:theme");
     theme_node.append_attribute("xmlns:a").set_value(constants::Namespaces.at("drawingml").c_str());
     theme_node.append_attribute("name").set_value("Office Theme");
     auto theme_elements_node = theme_node.append_child("a:themeElements");
     auto clr_scheme_node = theme_elements_node.append_child("a:clrScheme");
     clr_scheme_node.append_attribute("name").set_value("Office");
     
     struct scheme_element
     {
         std::string name;
         std::string sub_element_name;
         std::string val;
     };
     
     std::vector<scheme_element> scheme_elements =
     {
         {"a:dk1", "a:sysClr", "windowText"},
         {"a:lt1", "a:sysClr", "window"},
         {"a:dk2", "a:srgbClr", "1F497D"},
         {"a:lt2", "a:srgbClr", "EEECE1"},
         {"a:accent1", "a:srgbClr", "4F81BD"},
         {"a:accent2", "a:srgbClr", "C0504D"},
         {"a:accent3", "a:srgbClr", "9BBB59"},
         {"a:accent4", "a:srgbClr", "8064A2"},
         {"a:accent5", "a:srgbClr", "4BACC6"},
         {"a:accent6", "a:srgbClr", "F79646"},
         {"a:hlink", "a:srgbClr", "0000FF"},
         {"a:folHlink", "a:srgbClr", "800080"},
     };
     
     for(auto element : scheme_elements)
     {
         auto element_node = clr_scheme_node.append_child(element.name.c_str());
         element_node.append_child(element.sub_element_name.c_str()).append_attribute("val").set_value(element.val.c_str());

         if(element.name == "a:dk1")
         {
             element_node.child(element.sub_element_name.c_str()).append_attribute("lastClr").set_value("000000");
         }
         else if(element.name == "a:lt1")
         {
             element_node.child(element.sub_element_name.c_str()).append_attribute("lastClr").set_value("FFFFFF");
         }
     }

     struct font_scheme
     {
         bool typeface;
         std::string script;
         std::string major;
         std::string minor;
     };

     std::vector<font_scheme> font_schemes = 
     {
         {true, "a:latin", "Cambria", "Calibri"},
         {true, "a:ea", "", ""},
         {true, "a:cs", "", ""},
         {false, "Jpan", "\xef\xbc\xad\xef\xbc\xb3 \xef\xbc\xb0\xe3\x82\xb4\xe3\x82\xb7\xe3\x83\x83\xe3\x82\xaf", "\xef\xbc\xad\xef\xbc\xb3 \xef\xbc\xb0\xe3\x82\xb4\xe3\x82\xb7\xe3\x83\x83\xe3\x82\xaf"},
         {false, "Hang", "\xeb\xa7\x91\xec\x9d\x80 \xea\xb3\xa0\xeb\x94\x95", "\xeb\xa7\x91\xec\x9d\x80 \xea\xb3\xa0\xeb\x94\x95"},
         {false, "Hans", "\xe5\xae\x8b\xe4\xbd\x93", "\xe5\xae\x8b\xe4\xbd\x93"},
         {false, "Hant", "\xe6\x96\xb0\xe7\xb4\xb0\xe6\x98\x8e\xe9\xab\x94", "\xe6\x96\xb0\xe7\xb4\xb0\xe6\x98\x8e\xe9\xab\x94"},
         {false, "Arab", "Times New Roman", "Arial"},
         {false, "Hebr", "Times New Roman", "Arial"},
         {false, "Thai", "Tahoma", "Tahoma"},
         {false, "Ethi", "Nyala", "Nyala"},
         {false, "Beng", "Vrinda", "Vrinda"},
         {false, "Gujr", "Shruti", "Shruti"},
         {false, "Khmr", "MoolBoran", "DaunPenh"},
         {false, "Knda", "Tunga", "Tunga"},
         {false, "Guru", "Raavi", "Raavi"},
         {false, "Cans", "Euphemia", "Euphemia"},
         {false, "Cher", "Plantagenet Cherokee", "Plantagenet Cherokee"},
         {false, "Yiii", "Microsoft Yi Baiti", "Microsoft Yi Baiti"},
         {false, "Tibt", "Microsoft Himalaya", "Microsoft Himalaya"},
         {false, "Thaa", "MV Boli", "MV Boli"},
         {false, "Deva", "Mangal", "Mangal"},
         {false, "Telu", "Gautami", "Gautami"},
         {false, "Taml", "Latha", "Latha"},
         {false, "Syrc", "Estrangelo Edessa", "Estrangelo Edessa"},
         {false, "Orya", "Kalinga", "Kalinga"},
         {false, "Mlym", "Kartika", "Kartika"},
         {false, "Laoo", "DokChampa", "DokChampa"},
         {false, "Sinh", "Iskoola Pota", "Iskoola Pota"},
         {false, "Mong", "Mongolian Baiti", "Mongolian Baiti"},
         {false, "Viet", "Times New Roman", "Arial"},
         {false, "Uigh", "Microsoft Uighur", "Microsoft Uighur"}
     };

     auto font_scheme_node = theme_elements_node.append_child("a:fontScheme");
     font_scheme_node.append_attribute("name").set_value("Office");

     auto major_fonts_node = font_scheme_node.append_child("a:majorFont");
     auto minor_fonts_node = font_scheme_node.append_child("a:minorFont");

     for(auto scheme : font_schemes)
     {
         pugi::xml_node major_font_node, minor_font_node;

         if(scheme.typeface)
         {
             major_font_node = major_fonts_node.append_child(scheme.script.c_str());
             minor_font_node = minor_fonts_node.append_child(scheme.script.c_str());
         }
         else
         {
             major_font_node = major_fonts_node.append_child("a:font");
             major_font_node.append_attribute("script").set_value(scheme.script.c_str());
             minor_font_node = minor_fonts_node.append_child("a:font");
             minor_font_node.append_attribute("script").set_value(scheme.script.c_str());
         }

         major_font_node.append_attribute("typeface").set_value(scheme.major.c_str());
         minor_font_node.append_attribute("typeface").set_value(scheme.minor.c_str());
     }

     auto format_scheme_node = theme_elements_node.append_child("a:fmtScheme");
     format_scheme_node.append_attribute("name").set_value("Office");

     auto fill_style_list_node = format_scheme_node.append_child("a:fillStyleLst");
     fill_style_list_node.append_child("a:solidFill").append_child("a:schemeClr").append_attribute("val").set_value("phClr");

     auto grad_fill_node = fill_style_list_node.append_child("a:gradFill");
     grad_fill_node.append_attribute("rotWithShape").set_value(1);

     auto grad_fill_list = grad_fill_node.append_child("a:gsLst");
     auto gs_node = grad_fill_list.append_child("a:gs");
     gs_node.append_attribute("pos").set_value(0);
     auto scheme_color_node = gs_node.append_child("a:schemeClr");
     scheme_color_node.append_attribute("val").set_value("phClr");
     scheme_color_node.append_child("a:tint").append_attribute("val").set_value(50000);
     scheme_color_node.append_child("a:satMod").append_attribute("val").set_value(300000);

     gs_node = grad_fill_list.append_child("a:gs");
     gs_node.append_attribute("pos").set_value(35000);
     scheme_color_node = gs_node.append_child("a:schemeClr");
     scheme_color_node.append_attribute("val").set_value("phClr");
     scheme_color_node.append_child("a:tint").append_attribute("val").set_value(37000);
     scheme_color_node.append_child("a:satMod").append_attribute("val").set_value(300000);

     gs_node = grad_fill_list.append_child("a:gs");
     gs_node.append_attribute("pos").set_value(100000);
     scheme_color_node = gs_node.append_child("a:schemeClr");
     scheme_color_node.append_attribute("val").set_value("phClr");
     scheme_color_node.append_child("a:tint").append_attribute("val").set_value(15000);
     scheme_color_node.append_child("a:satMod").append_attribute("val").set_value(350000);

     auto lin_node = grad_fill_node.append_child("a:lin");
     lin_node.append_attribute("ang").set_value(16200000);
     lin_node.append_attribute("scaled").set_value(1);

     grad_fill_node = fill_style_list_node.append_child("a:gradFill");
     grad_fill_node.append_attribute("rotWithShape").set_value(1);

     grad_fill_list = grad_fill_node.append_child("a:gsLst");
     gs_node = grad_fill_list.append_child("a:gs");
     gs_node.append_attribute("pos").set_value(0);
     scheme_color_node = gs_node.append_child("a:schemeClr");
     scheme_color_node.append_attribute("val").set_value("phClr");
     scheme_color_node.append_child("a:shade").append_attribute("val").set_value(51000);
     scheme_color_node.append_child("a:satMod").append_attribute("val").set_value(130000);

     gs_node = grad_fill_list.append_child("a:gs");
     gs_node.append_attribute("pos").set_value(80000);
     scheme_color_node = gs_node.append_child("a:schemeClr");
     scheme_color_node.append_attribute("val").set_value("phClr");
     scheme_color_node.append_child("a:shade").append_attribute("val").set_value(93000);
     scheme_color_node.append_child("a:satMod").append_attribute("val").set_value(130000);

     gs_node = grad_fill_list.append_child("a:gs");
     gs_node.append_attribute("pos").set_value(100000);
     scheme_color_node = gs_node.append_child("a:schemeClr");
     scheme_color_node.append_attribute("val").set_value("phClr");
     scheme_color_node.append_child("a:shade").append_attribute("val").set_value(94000);
     scheme_color_node.append_child("a:satMod").append_attribute("val").set_value(135000);

     lin_node = grad_fill_node.append_child("a:lin");
     lin_node.append_attribute("ang").set_value(16200000);
     lin_node.append_attribute("scaled").set_value(0);

     auto line_style_list_node = format_scheme_node.append_child("a:lnStyleLst");

     auto ln_node = line_style_list_node.append_child("a:ln");
     ln_node.append_attribute("w").set_value(9525);
     ln_node.append_attribute("cap").set_value("flat");
     ln_node.append_attribute("cmpd").set_value("sng");
     ln_node.append_attribute("algn").set_value("ctr");

     auto solid_fill_node = ln_node.append_child("a:solidFill");
     scheme_color_node = solid_fill_node.append_child("a:schemeClr");
     scheme_color_node.append_attribute("val").set_value("phClr");
     scheme_color_node.append_child("a:shade").append_attribute("val").set_value(95000);
     scheme_color_node.append_child("a:satMod").append_attribute("val").set_value(105000);
     ln_node.append_child("a:prstDash").append_attribute("val").set_value("solid");

     ln_node = line_style_list_node.append_child("a:ln");
     ln_node.append_attribute("w").set_value(25400);
     ln_node.append_attribute("cap").set_value("flat");
     ln_node.append_attribute("cmpd").set_value("sng");
     ln_node.append_attribute("algn").set_value("ctr");

     solid_fill_node = ln_node.append_child("a:solidFill");
     scheme_color_node = solid_fill_node.append_child("a:schemeClr");
     scheme_color_node.append_attribute("val").set_value("phClr");
     ln_node.append_child("a:prstDash").append_attribute("val").set_value("solid");

     ln_node = line_style_list_node.append_child("a:ln");
     ln_node.append_attribute("w").set_value(38100);
     ln_node.append_attribute("cap").set_value("flat");
     ln_node.append_attribute("cmpd").set_value("sng");
     ln_node.append_attribute("algn").set_value("ctr");

     solid_fill_node = ln_node.append_child("a:solidFill");
     scheme_color_node = solid_fill_node.append_child("a:schemeClr");
     scheme_color_node.append_attribute("val").set_value("phClr");
     ln_node.append_child("a:prstDash").append_attribute("val").set_value("solid");

     auto effect_style_list_node = format_scheme_node.append_child("a:effectStyleLst");
     auto effect_style_node = effect_style_list_node.append_child("a:effectStyle");
     auto effect_list_node = effect_style_node.append_child("a:effectLst");
     auto outer_shadow_node = effect_list_node.append_child("a:outerShdw");
     outer_shadow_node.append_attribute("blurRad").set_value(40000);
     outer_shadow_node.append_attribute("dist").set_value(20000);
     outer_shadow_node.append_attribute("dir").set_value(5400000);
     outer_shadow_node.append_attribute("rotWithShape").set_value(0);
     auto srgb_clr_node = outer_shadow_node.append_child("a:srgbClr");
     srgb_clr_node.append_attribute("val").set_value("000000");
     srgb_clr_node.append_child("a:alpha").append_attribute("val").set_value(38000);

     effect_style_node = effect_style_list_node.append_child("a:effectStyle");
     effect_list_node = effect_style_node.append_child("a:effectLst");
     outer_shadow_node = effect_list_node.append_child("a:outerShdw");
     outer_shadow_node.append_attribute("blurRad").set_value(40000);
     outer_shadow_node.append_attribute("dist").set_value(23000);
     outer_shadow_node.append_attribute("dir").set_value(5400000);
     outer_shadow_node.append_attribute("rotWithShape").set_value(0);
     srgb_clr_node = outer_shadow_node.append_child("a:srgbClr");
     srgb_clr_node.append_attribute("val").set_value("000000");
     srgb_clr_node.append_child("a:alpha").append_attribute("val").set_value(35000);

     effect_style_node = effect_style_list_node.append_child("a:effectStyle");
     effect_list_node = effect_style_node.append_child("a:effectLst");
     outer_shadow_node = effect_list_node.append_child("a:outerShdw");
     outer_shadow_node.append_attribute("blurRad").set_value(40000);
     outer_shadow_node.append_attribute("dist").set_value(23000);
     outer_shadow_node.append_attribute("dir").set_value(5400000);
     outer_shadow_node.append_attribute("rotWithShape").set_value(0);
     srgb_clr_node = outer_shadow_node.append_child("a:srgbClr");
     srgb_clr_node.append_attribute("val").set_value("000000");
     srgb_clr_node.append_child("a:alpha").append_attribute("val").set_value(35000);
     auto scene3d_node = effect_style_node.append_child("a:scene3d");
     auto camera_node = scene3d_node.append_child("a:camera");
     camera_node.append_attribute("prst").set_value("orthographicFront");
     auto rot_node = camera_node.append_child("a:rot");
     rot_node.append_attribute("lat").set_value(0);
     rot_node.append_attribute("lon").set_value(0);
     rot_node.append_attribute("rev").set_value(0);
     auto light_rig_node = scene3d_node.append_child("a:lightRig");
     light_rig_node.append_attribute("rig").set_value("threePt");
     light_rig_node.append_attribute("dir").set_value("t");
     rot_node = light_rig_node.append_child("a:rot");
     rot_node.append_attribute("lat").set_value(0);
     rot_node.append_attribute("lon").set_value(0);
     rot_node.append_attribute("rev").set_value(1200000);

     auto bevel_node = effect_style_node.append_child("a:sp3d").append_child("a:bevelT");
     bevel_node.append_attribute("w").set_value(63500);
     bevel_node.append_attribute("h").set_value(25400);

     auto bg_fill_style_list_node = format_scheme_node.append_child("a:bgFillStyleLst");

     bg_fill_style_list_node.append_child("a:solidFill").append_child("a:schemeClr").append_attribute("val").set_value("phClr");

     grad_fill_node = bg_fill_style_list_node.append_child("a:gradFill");
     grad_fill_node.append_attribute("rotWithShape").set_value(1);

     grad_fill_list = grad_fill_node.append_child("a:gsLst");
     gs_node = grad_fill_list.append_child("a:gs");
     gs_node.append_attribute("pos").set_value(0);
     scheme_color_node = gs_node.append_child("a:schemeClr");
     scheme_color_node.append_attribute("val").set_value("phClr");
     scheme_color_node.append_child("a:tint").append_attribute("val").set_value(40000);
     scheme_color_node.append_child("a:satMod").append_attribute("val").set_value(350000);

     gs_node = grad_fill_list.append_child("a:gs");
     gs_node.append_attribute("pos").set_value(40000);
     scheme_color_node = gs_node.append_child("a:schemeClr");
     scheme_color_node.append_attribute("val").set_value("phClr");
     scheme_color_node.append_child("a:tint").append_attribute("val").set_value(45000);
     scheme_color_node.append_child("a:shade").append_attribute("val").set_value(99000);
     scheme_color_node.append_child("a:satMod").append_attribute("val").set_value(350000);

     gs_node = grad_fill_list.append_child("a:gs");
     gs_node.append_attribute("pos").set_value(100000);
     scheme_color_node = gs_node.append_child("a:schemeClr");
     scheme_color_node.append_attribute("val").set_value("phClr");
     scheme_color_node.append_child("a:shade").append_attribute("val").set_value(20000);
     scheme_color_node.append_child("a:satMod").append_attribute("val").set_value(255000);

     auto path_node = grad_fill_node.append_child("a:path");
     path_node.append_attribute("path").set_value("circle");
     auto fill_to_rect_node = path_node.append_child("a:fillToRect");
     fill_to_rect_node.append_attribute("l").set_value(50000);
     fill_to_rect_node.append_attribute("t").set_value(-80000);
     fill_to_rect_node.append_attribute("r").set_value(50000);
     fill_to_rect_node.append_attribute("b").set_value(180000);

     grad_fill_node = bg_fill_style_list_node.append_child("a:gradFill");
     grad_fill_node.append_attribute("rotWithShape").set_value(1);

     grad_fill_list = grad_fill_node.append_child("a:gsLst");
     gs_node = grad_fill_list.append_child("a:gs");
     gs_node.append_attribute("pos").set_value(0);
     scheme_color_node = gs_node.append_child("a:schemeClr");
     scheme_color_node.append_attribute("val").set_value("phClr");
     scheme_color_node.append_child("a:tint").append_attribute("val").set_value(80000);
     scheme_color_node.append_child("a:satMod").append_attribute("val").set_value(300000);

     gs_node = grad_fill_list.append_child("a:gs");
     gs_node.append_attribute("pos").set_value(100000);
     scheme_color_node = gs_node.append_child("a:schemeClr");
     scheme_color_node.append_attribute("val").set_value("phClr");
     scheme_color_node.append_child("a:shade").append_attribute("val").set_value(30000);
     scheme_color_node.append_child("a:satMod").append_attribute("val").set_value(200000);

     path_node = grad_fill_node.append_child("a:path");
     path_node.append_attribute("path").set_value("circle");
     fill_to_rect_node = path_node.append_child("a:fillToRect");
     fill_to_rect_node.append_attribute("l").set_value(50000);
     fill_to_rect_node.append_attribute("t").set_value(50000);
     fill_to_rect_node.append_attribute("r").set_value(50000);
     fill_to_rect_node.append_attribute("b").set_value(50000);

     theme_node.append_child("a:objectDefaults");
     theme_node.append_child("a:extraClrSchemeLst");
     
     std::stringstream ss;
     doc.print(ss);
     return ss.str();
}

}
