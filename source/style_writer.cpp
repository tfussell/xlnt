#include <sstream>
#include <pugixml.hpp>

#include "writer/style_writer.hpp"
#include "workbook/workbook.hpp"
#include "worksheet/worksheet.hpp"
#include "worksheet/range.hpp"
#include "cell/cell.hpp"

namespace xlnt {

style_writer::style_writer(xlnt::workbook &wb) : wb_(wb)
{
  
}

std::unordered_map<std::size_t, std::string> style_writer::get_style_by_hash() const
{
    std::unordered_map<std::size_t, std::string> styles;
    for(auto ws : wb_)
    {
        for(auto row : ws.rows())
        {
            for(auto cell : row)
            {
                if(cell.has_style())
                {
                    styles[1] = "style";
                }
            }
        }
    }
    return styles;
}

std::vector<style> style_writer::get_styles() const
{
    std::vector<style> styles;
    
    for(auto ws : wb_)
    {
        for(auto row : ws.rows())
        {
            for(auto cell : row)
            {
                if(cell.has_style())
                {
                    styles.push_back(cell.get_style());
                }
            }
        }
    }
    
    return styles;
}
    
std::string style_writer::write_table() const
{
    pugi::xml_document doc;
    auto style_sheet_node = doc.append_child("styleSheet");
    style_sheet_node.append_attribute("xmlns").set_value("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    style_sheet_node.append_attribute("xmlns:mc").set_value("http://schemas.openxmlformats.org/markup-compatibility/2006");
    style_sheet_node.append_attribute("mc:Ignorable").set_value("x14ac");
    style_sheet_node.append_attribute("xmlns:x14ac").set_value("http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

    auto fonts_node = style_sheet_node.append_child("fonts");
    fonts_node.append_attribute("count").set_value(1);
    fonts_node.append_attribute("x14ac:knownFonts").set_value(1);

    auto font_node = fonts_node.append_child("font");
    auto size_node = font_node.append_child("sz");
    size_node.append_attribute("val").set_value(11);
    auto color_node = font_node.append_child("color");
    color_node.append_attribute("theme").set_value(1);
    auto name_node = font_node.append_child("name");
    name_node.append_attribute("val").set_value("Calibri");
    auto family_node = font_node.append_child("family");
    family_node.append_attribute("val").set_value(2);
    auto scheme_node = font_node.append_child("scheme");
    scheme_node.append_attribute("val").set_value("minor");

    auto fills_node = style_sheet_node.append_child("fills");
    fills_node.append_attribute("count").set_value(2);
    fills_node.append_child("fill").append_child("patternFill").append_attribute("patternType").set_value("none");
    fills_node.append_child("fill").append_child("patternFill").append_attribute("patternType").set_value("gray125");

    auto borders_node = style_sheet_node.append_child("borders");
    borders_node.append_attribute("count").set_value(1);
    auto border_node = borders_node.append_child("border");
    border_node.append_child("left");
    border_node.append_child("right");
    border_node.append_child("top");
    border_node.append_child("bottom");
    border_node.append_child("diagonal");

    auto cell_style_xfs_node = style_sheet_node.append_child("cellStyleXfs");
    cell_style_xfs_node.append_attribute("count").set_value(1);
    auto xf_node = cell_style_xfs_node.append_child("xf");
    xf_node.append_attribute("numFmtId").set_value(0);
    xf_node.append_attribute("fontId").set_value(0);
    xf_node.append_attribute("fillId").set_value(0);
    xf_node.append_attribute("borderId").set_value(0);

    auto cell_xfs_node = style_sheet_node.append_child("cellXfs");
    cell_xfs_node.append_attribute("count").set_value(1);
    xf_node = cell_xfs_node.append_child("xf");
    xf_node.append_attribute("numFmtId").set_value(0);
    xf_node.append_attribute("fontId").set_value(0);
    xf_node.append_attribute("fillId").set_value(0);
    xf_node.append_attribute("borderId").set_value(0);
    xf_node.append_attribute("xfId").set_value(0);

    auto cell_styles_node = style_sheet_node.append_child("cellStyles");
    cell_styles_node.append_attribute("count").set_value(1);
    auto cell_style_node = cell_styles_node.append_child("cellStyle");
    cell_style_node.append_attribute("name").set_value("Normal");
    cell_style_node.append_attribute("xfId").set_value(0);
    cell_style_node.append_attribute("builtinId").set_value(0);

    style_sheet_node.append_child("dxfs").append_attribute("count").set_value(0);

    auto table_styles_node = style_sheet_node.append_child("tableStyles");
    table_styles_node.append_attribute("count").set_value(0);
    table_styles_node.append_attribute("defaultTableStyle").set_value("TableStyleMedium2");
    table_styles_node.append_attribute("defaultPivotStyle").set_value("PivotStyleMedium9");

    auto ext_list_node = style_sheet_node.append_child("extLst");
    auto ext_node = ext_list_node.append_child("ext");
    ext_node.append_attribute("uri").set_value("{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}");
    ext_node.append_attribute("xmlns:x14").set_value("http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
    ext_node.append_child("x14:slicerStyles").append_attribute("defaultSlicerStyle").set_value("SlicerStyleLight1");

    std::stringstream ss;
    doc.save(ss);

    return ss.str();
}
    
} // namespace xlnt
