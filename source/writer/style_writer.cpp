#include <sstream>
#include <pugixml.hpp>

#include <xlnt/styles/alignment.hpp>
#include <xlnt/styles/border.hpp>
#include <xlnt/styles/fill.hpp>
#include <xlnt/styles/font.hpp>
#include <xlnt/styles/number_format.hpp>
#include <xlnt/styles/protection.hpp>
#include <xlnt/writer/style_writer.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/worksheet/worksheet.hpp>
#include <xlnt/worksheet/range.hpp>
#include <xlnt/cell/cell.hpp>

namespace xlnt {

style_writer::style_writer(xlnt::workbook &wb) : wb_(wb)
{
  
}

std::string style_writer::write_table() const
{
    pugi::xml_document doc;
    auto style_sheet_node = doc.append_child("styleSheet");
    style_sheet_node.append_attribute("xmlns").set_value("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    style_sheet_node.append_attribute("xmlns:mc").set_value("http://schemas.openxmlformats.org/markup-compatibility/2006");
    style_sheet_node.append_attribute("mc:Ignorable").set_value("x14ac");
    style_sheet_node.append_attribute("xmlns:x14ac").set_value("http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
    
    auto num_fmts_node = style_sheet_node.append_child("numFmts");
    auto num_fmts = wb_.get_number_formats();
    num_fmts_node.append_attribute("count").set_value(static_cast<int>(num_fmts.size()));
    
    int custom_index = 164;
    
    for(auto &num_fmt : num_fmts)
    {
        auto num_fmt_node = num_fmts_node.append_child("numFmt");
        
        if(num_fmt.get_format_code() == number_format::format::unknown)
        {
            num_fmt_node.append_attribute("numFmtId").set_value(custom_index++);
        }
        else
        {
            num_fmt_node.append_attribute("numFmtId").set_value(num_fmt.get_format_index());
        }
        
        num_fmt_node.append_attribute("formatCode").set_value(num_fmt.get_format_string().c_str());
    }

    auto fonts_node = style_sheet_node.append_child("fonts");
    auto fonts = wb_.get_fonts();
    if(fonts.empty())
    {
        fonts.push_back(font());
    }
    fonts_node.append_attribute("count").set_value(static_cast<int>(fonts.size()));
    fonts_node.append_attribute("x14ac:knownFonts").set_value(1);
    
    for(auto &f : fonts)
    {
        auto font_node = fonts_node.append_child("font");
        auto size_node = font_node.append_child("sz");
        size_node.append_attribute("val").set_value(f.get_size());
        auto color_node = font_node.append_child("color");
        color_node.append_attribute("theme").set_value(1);
        auto name_node = font_node.append_child("name");
        name_node.append_attribute("val").set_value(f.get_name().c_str());
        auto family_node = font_node.append_child("family");
        family_node.append_attribute("val").set_value(2);
        auto scheme_node = font_node.append_child("scheme");
        scheme_node.append_attribute("val").set_value("minor");
        
        if(f.is_bold())
        {
            auto bold_node = font_node.append_child("b");
            bold_node.append_attribute("val").set_value(1);
        }
    }

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
    const auto &styles = wb_.get_styles();
    cell_xfs_node.append_attribute("count").set_value(static_cast<int>(styles.size()));
    
    custom_index = 164;
    
    for(auto &style : styles)
    {
        xf_node = cell_xfs_node.append_child("xf");
        if(style.get_number_format().get_format_code() == number_format::format::unknown)
        {
            xf_node.append_attribute("numFmtId").set_value(custom_index++);
        }
        else
        {
            xf_node.append_attribute("numFmtId").set_value(style.get_number_format().get_format_index());
        }
        xf_node.append_attribute("applyNumberFormat").set_value(1);
        xf_node.append_attribute("fontId").set_value((int)style.get_font_index());
        xf_node.append_attribute("fillId").set_value((int)style.get_fill_index());
        xf_node.append_attribute("borderId").set_value((int)style.get_border_index());
        
        auto alignment_node = xf_node.append_child("alignment");
        alignment_node.append_attribute("vertical").set_value("top");
        alignment_node.append_attribute("wrapText").set_value(1);
    }

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

std::string style_writer::write_number_formats()
{
    pugi::xml_document doc;

    auto root = doc.append_child("styleSheet");
    root.append_attribute("xmlns").set_value("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    
    auto num_fmts_node = root.append_child("numFmts");
    num_fmts_node.append_attribute("count").set_value(1);
    
    auto num_fmt_node = num_fmts_node.append_child("numFmt");
    num_fmt_node.append_attribute("formatCode").set_value("YYYY");
    num_fmt_node.append_attribute("numFmtId").set_value(164);
    
    std::stringstream ss;
    doc.save(ss);
    
    return ss.str();
}
    
} // namespace xlnt
