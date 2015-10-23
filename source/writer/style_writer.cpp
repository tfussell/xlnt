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
    
    for(auto &num_fmt : num_fmts)
    {
        auto num_fmt_node = num_fmts_node.append_child("numFmt");
        num_fmt_node.append_attribute("numFmtId").set_value(num_fmt.get_id());
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
        
        if(f.get_color().get_type() == color::type::indexed)
        {
            color_node.append_attribute("indexed").set_value(static_cast<unsigned int>(f.get_color().get_index()));
        }
        else if(f.get_color().get_type() == color::type::theme)
        {
            color_node.append_attribute("theme").set_value(static_cast<unsigned int>(f.get_color().get_theme()));
        }
        
        auto name_node = font_node.append_child("name");
        name_node.append_attribute("val").set_value(f.get_name().c_str());
        
        if(f.has_family())
        {
            auto family_node = font_node.append_child("family");
            family_node.append_attribute("val").set_value(2);
        }
        
        if(f.has_scheme())
        {
            auto scheme_node = font_node.append_child("scheme");
            scheme_node.append_attribute("val").set_value("minor");
        }
        
        if(f.is_bold())
        {
            auto bold_node = font_node.append_child("b");
            bold_node.append_attribute("val").set_value(1);
        }
    }

    auto fills_node = style_sheet_node.append_child("fills");
    const auto &fills = wb_.get_fills();
    fills_node.append_attribute("count").set_value(static_cast<unsigned int>(fills.size()));
    
    for(auto &fill_ : fills)
    {
        auto fill_node = fills_node.append_child("fill");
        
        if(fill_.get_type() == fill::type::pattern)
        {
            auto pattern_fill_node = fill_node.append_child("patternFill");
            auto type_string = fill_.get_pattern_type_string();
            pattern_fill_node.append_attribute("patternType").set_value(type_string.c_str());
            
            if(fill_.has_foreground_color())
            {
                auto fg_color_node = pattern_fill_node.append_child("fgColor");
                
                switch(fill_.get_foreground_color().get_type())
                {
                    case color::type::auto_: fg_color_node.append_attribute("auto").set_value(fill_.get_foreground_color().get_auto()); break;
                    case color::type::theme: fg_color_node.append_attribute("theme").set_value(fill_.get_foreground_color().get_theme()); break;
                    case color::type::indexed: fg_color_node.append_attribute("indexed").set_value(fill_.get_foreground_color().get_index()); break;
                    default: throw std::runtime_error("bad type");
                }
            }
            
            if(fill_.has_background_color())
            {
                auto bg_color_node = pattern_fill_node.append_child("bgColor");
                
                switch(fill_.get_background_color().get_type())
                {
                    case color::type::auto_: bg_color_node.append_attribute("auto").set_value(fill_.get_background_color().get_auto()); break;
                    case color::type::theme: bg_color_node.append_attribute("theme").set_value(fill_.get_background_color().get_theme()); break;
                    case color::type::indexed: bg_color_node.append_attribute("indexed").set_value(fill_.get_background_color().get_index()); break;
                    default: throw std::runtime_error("bad type");
                }
            }
        }
    }

    auto borders_node = style_sheet_node.append_child("borders");
    const auto &borders = wb_.get_borders();
    borders_node.append_attribute("count").set_value(static_cast<unsigned int>(borders.size()));
    
    for(const auto &border_ : borders)
    {
        auto border_node = borders_node.append_child("border");
        
        if(border_.left_assigned)
        {
            auto left_node = border_node.append_child("left");
            
            if(border_.left.is_style_assigned())
            {
                auto style_attribute = left_node.append_attribute("style");
                
                switch(border_.left.get_style())
                {
                    case border_style::none: style_attribute.set_value("none"); break;
                    case border_style::thin: style_attribute.set_value("thin"); break;
                    default: throw std::runtime_error("invalid style");
                }
            }
            
            if(border_.left.is_color_assigned())
            {
                auto color_node = left_node.append_child("color");
                
                if(border_.left.get_color_type() == side::color_type::indexed)
                {
                    color_node.append_attribute("indexed").set_value((int)border_.left.get_color());
                }
                else if(border_.diagonal.get_color_type() == side::color_type::theme)
                {
                    color_node.append_attribute("indexed").set_value((int)border_.left.get_color());
                }
                else
                {
                    throw std::runtime_error("invalid color type");
                }
            }
        }
        
        if(border_.right_assigned)
        {
            auto right_node = border_node.append_child("right");
            
            if(border_.right.is_style_assigned())
            {
                auto style_attribute = right_node.append_attribute("style");
                
                switch(border_.right.get_style())
                {
                    case border_style::none: style_attribute.set_value("none"); break;
                    case border_style::thin: style_attribute.set_value("thin"); break;
                    default: throw std::runtime_error("invalid style");
                }
            }
            
            if(border_.right.is_color_assigned())
            {
                auto color_node = right_node.append_child("color");
                
                if(border_.right.get_color_type() == side::color_type::indexed)
                {
                    color_node.append_attribute("indexed").set_value((int)border_.right.get_color());
                }
                else if(border_.diagonal.get_color_type() == side::color_type::theme)
                {
                    color_node.append_attribute("indexed").set_value((int)border_.right.get_color());
                }
                else
                {
                    throw std::runtime_error("invalid color type");
                }
            }
        }
        
        if(border_.top_assigned)
        {
            auto top_node = border_node.append_child("top");
            
            if(border_.top.is_style_assigned())
            {
                auto style_attribute = top_node.append_attribute("style");
                
                switch(border_.top.get_style())
                {
                    case border_style::none: style_attribute.set_value("none"); break;
                    case border_style::thin: style_attribute.set_value("thin"); break;
                    default: throw std::runtime_error("invalid style");
                }
            }
            
            if(border_.top.is_color_assigned())
            {
                auto color_node = top_node.append_child("color");
                if(border_.top.get_color_type() == side::color_type::indexed)
                {
                    color_node.append_attribute("indexed").set_value((int)border_.top.get_color());
                }
                else if(border_.diagonal.get_color_type() == side::color_type::theme)
                {
                    color_node.append_attribute("indexed").set_value((int)border_.top.get_color());
                }
                else
                {
                    throw std::runtime_error("invalid color type");
                }
            }
        }
        
        if(border_.bottom_assigned)
        {
            auto bottom_node = border_node.append_child("bottom");
            
            if(border_.bottom.is_style_assigned())
            {
                auto style_attribute = bottom_node.append_attribute("style");
                
                switch(border_.bottom.get_style())
                {
                    case border_style::none: style_attribute.set_value("none"); break;
                    case border_style::thin: style_attribute.set_value("thin"); break;
                    default: throw std::runtime_error("invalid style");
                }
            }
            
            if(border_.bottom.is_color_assigned())
            {
                auto color_node = bottom_node.append_child("color");
                
                if(border_.bottom.get_color_type() == side::color_type::indexed)
                {
                    color_node.append_attribute("indexed").set_value((int)border_.bottom.get_color());
                }
                else if(border_.diagonal.get_color_type() == side::color_type::theme)
                {
                    color_node.append_attribute("indexed").set_value((int)border_.bottom.get_color());
                }
                else
                {
                    throw std::runtime_error("invalid color type");
                }
            }
        }
        
        if(border_.diagonal_assigned)
        {
            auto diagonal_node = border_node.append_child("diagonal");
            
            if(border_.diagonal.is_style_assigned())
            {
                auto style_attribute = diagonal_node.append_attribute("style");
                
                switch(border_.diagonal.get_style())
                {
                    case border_style::none: style_attribute.set_value("none"); break;
                    case border_style::thin: style_attribute.set_value("thin"); break;
                    default: throw std::runtime_error("invalid style");
                }
            }
            
            if(border_.diagonal.is_color_assigned())
            {
                auto color_node = diagonal_node.append_child("color");
                
                if(border_.diagonal.get_color_type() == side::color_type::indexed)
                {
                    color_node.append_attribute("indexed").set_value((int)border_.diagonal.get_color());
                }
                else if(border_.diagonal.get_color_type() == side::color_type::theme)
                {
                    color_node.append_attribute("indexed").set_value((int)border_.diagonal.get_color());
                }
                else
                {
                    throw std::runtime_error("invalid color type");
                }
            }
        }
    }

    auto cell_style_xfs_node = style_sheet_node.append_child("cellStyleXfs");
    cell_style_xfs_node.append_attribute("count").set_value(1);
    auto style_xf_node = cell_style_xfs_node.append_child("xf");
    style_xf_node.append_attribute("numFmtId").set_value(0);
    style_xf_node.append_attribute("fontId").set_value(0);
    style_xf_node.append_attribute("fillId").set_value(0);
    style_xf_node.append_attribute("borderId").set_value(0);
    
    auto cell_xfs_node = style_sheet_node.append_child("cellXfs");
    const auto &styles = wb_.get_styles();
    cell_xfs_node.append_attribute("count").set_value(static_cast<int>(styles.size()));
    
    for(auto &style : styles)
    {
        auto xf_node = cell_xfs_node.append_child("xf");
        xf_node.append_attribute("numFmtId").set_value(style.get_number_format().get_id());
        xf_node.append_attribute("fontId").set_value((int)style.get_font_id());
        
        if(style.fill_apply_)
        {
            xf_node.append_attribute("fillId").set_value((int)style.get_fill_id());
        }
        
        if(style.border_apply_)
        {
            xf_node.append_attribute("borderId").set_value((int)style.get_border_id());
        }
        
        xf_node.append_attribute("applyNumberFormat").set_value(style.number_format_apply_ ? 1 : 0);
        xf_node.append_attribute("applyFont").set_value(style.font_apply_ ? 1 : 0);
        xf_node.append_attribute("applyFill").set_value(style.fill_apply_ ? 1 : 0);
        xf_node.append_attribute("applyBorder").set_value(style.border_apply_ ? 1 : 0);
        xf_node.append_attribute("applyAlignment").set_value(style.alignment_apply_ ? 1 : 0);
        xf_node.append_attribute("applyProtection").set_value(style.protection_apply_ ? 1 : 0);
        
        if(style.alignment_apply_)
        {
            auto alignment_node = xf_node.append_child("alignment");
            
            if(style.alignment_.has_vertical())
            {
                switch(style.alignment_.get_vertical())
                {
                    case alignment::vertical_alignment::bottom:
                        alignment_node.append_attribute("vertical").set_value("bottom");
                        break;
                    case alignment::vertical_alignment::center:
                        alignment_node.append_attribute("vertical").set_value("center");
                        break;
                    case alignment::vertical_alignment::justify:
                        alignment_node.append_attribute("vertical").set_value("justify");
                        break;
                    case alignment::vertical_alignment::top:
                        alignment_node.append_attribute("vertical").set_value("top");
                        break;
                    default:
                        throw std::runtime_error("invalid alignment");
                }
            }
            
            if(style.alignment_.has_horizontal())
            {
                switch(style.alignment_.get_horizontal())
                {
                    case alignment::horizontal_alignment::center:
                        alignment_node.append_attribute("horizontal").set_value("center");
                        break;
                    case alignment::horizontal_alignment::center_continuous:
                        alignment_node.append_attribute("horizontal").set_value("center_continuous");
                        break;
                    case alignment::horizontal_alignment::general:
                        alignment_node.append_attribute("horizontal").set_value("general");
                        break;
                    case alignment::horizontal_alignment::justify:
                        alignment_node.append_attribute("horizontal").set_value("justify");
                        break;
                    case alignment::horizontal_alignment::left:
                        alignment_node.append_attribute("horizontal").set_value("left");
                        break;
                    case alignment::horizontal_alignment::right:
                        alignment_node.append_attribute("horizontal").set_value("right");
                        break;
                    default:
                        throw std::runtime_error("invalid alignment");
                }
            }
            
            if(style.alignment_.get_wrap_text())
            {
                alignment_node.append_attribute("wrapText").set_value(1);
            }
        }
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
    
    auto colors_node = style_sheet_node.append_child("colors");
    auto indexed_colors_node = colors_node.append_child("indexedColors");
    
    const std::vector<std::string> colors =
    {
        "ff000000",
        "ffffffff",
        "ffff0000",
        "ff00ff00",
        "ff0000ff",
        "ffffff00",
        "ffff00ff",
        "ff00ffff",
        "ff000000",
        "ffaaaaaa",
        "ffbdc0bf",
        "ffdbdbdb"
    };
    
    for(auto &color : colors)
    {
        indexed_colors_node.append_child("rgbColor").append_attribute("rgb").set_value(color.c_str());
    }

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
