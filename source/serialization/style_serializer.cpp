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
#include <cctype> // for std::tolower
#include <unordered_set>

#include <xlnt/serialization/style_serializer.hpp>
#include <xlnt/cell/cell.hpp>
#include <xlnt/serialization/xml_document.hpp>
#include <xlnt/serialization/xml_node.hpp>
#include <xlnt/styles/alignment.hpp>
#include <xlnt/styles/border.hpp>
#include <xlnt/styles/cell_style.hpp>
#include <xlnt/styles/fill.hpp>
#include <xlnt/styles/font.hpp>
#include <xlnt/styles/named_style.hpp>
#include <xlnt/styles/number_format.hpp>
#include <xlnt/styles/protection.hpp>
#include <xlnt/workbook/worksheet_iterator.hpp>
#include <xlnt/worksheet/cell_iterator.hpp>
#include <xlnt/worksheet/cell_vector.hpp>
#include <xlnt/worksheet/worksheet.hpp>
#include <xlnt/worksheet/range_iterator.hpp>

namespace {

bool equals_case_insensitive(const std::string &left, const std::string &right)
{
    if (left.size() != right.size())
    {
        return false;
    }

    for (std::size_t i = 0; i < left.size(); i++)
    {
        if (std::tolower(left[i]) != std::tolower(right[i]))
        {
            return false;
        }
    }
    
    return true;
}

bool is_true(const std::string &bool_string)
{
    return bool_string == "1";
}

xlnt::protection::type protection_type_from_string(const std::string &type_string)
{
    if (equals_case_insensitive(type_string, "true"))
    {
        return xlnt::protection::type::protected_;
    }
    else if (equals_case_insensitive(type_string, "inherit"))
    {
        return xlnt::protection::type::inherit;
    }

    return xlnt::protection::type::unprotected;
};

xlnt::font::underline_style underline_style_from_string(const std::string &underline_string)
{
    if (equals_case_insensitive(underline_string, "none"))
    {
        return xlnt::font::underline_style::none;
    }
    else if (equals_case_insensitive(underline_string, "single"))
    {
        return xlnt::font::underline_style::single;
    }
    else if (equals_case_insensitive(underline_string, "single-accounting"))
    {
        return xlnt::font::underline_style::single_accounting;
    }
    else if (equals_case_insensitive(underline_string, "double"))
    {
        return xlnt::font::underline_style::double_;
    }
    else if (equals_case_insensitive(underline_string, "double-accounting"))
    {
        return xlnt::font::underline_style::double_accounting;
    }

    return xlnt::font::underline_style::none;
}

xlnt::fill::pattern_type pattern_fill_type_from_string(const std::string &fill_type)
{
    if (equals_case_insensitive(fill_type, "none")) return xlnt::fill::pattern_type::none;
    if (equals_case_insensitive(fill_type, "solid")) return xlnt::fill::pattern_type::solid;
    if (equals_case_insensitive(fill_type, "gray125")) return xlnt::fill::pattern_type::gray125;
    return xlnt::fill::pattern_type::none;
};

xlnt::border_style border_style_from_string(const std::string &border_style_string)
{
    if (equals_case_insensitive(border_style_string, "none"))
    {
        return xlnt::border_style::none;
    }
    else if (equals_case_insensitive(border_style_string, "dashdot"))
    {
        return xlnt::border_style::dashdot;
    }
    else if (equals_case_insensitive(border_style_string, "dashdotdot"))
    {
        return xlnt::border_style::dashdotdot;
    }
    else if (equals_case_insensitive(border_style_string, "dashed"))
    {
        return xlnt::border_style::dashed;
    }
    else if (equals_case_insensitive(border_style_string, "dotted"))
    {
        return xlnt::border_style::dotted;
    }
    else if (equals_case_insensitive(border_style_string, "double"))
    {
        return xlnt::border_style::double_;
    }
    else if (equals_case_insensitive(border_style_string, "hair"))
    {
        return xlnt::border_style::hair;
    }
    else if (equals_case_insensitive(border_style_string, "medium"))
    {
        return xlnt::border_style::medium;
    }
    else if (equals_case_insensitive(border_style_string, "mediumdashdot"))
    {
        return xlnt::border_style::mediumdashdot;
    }
    else if (equals_case_insensitive(border_style_string, "mediumdashdotdot"))
    {
        return xlnt::border_style::mediumdashdotdot;
    }
    else if (equals_case_insensitive(border_style_string, "mediumdashed"))
    {
        return xlnt::border_style::mediumdashed;
    }
    else if (equals_case_insensitive(border_style_string, "slantdashdot"))
    {
        return xlnt::border_style::slantdashdot;
    }
    else if (equals_case_insensitive(border_style_string, "thick"))
    {
        return xlnt::border_style::thick;
    }
    else if (equals_case_insensitive(border_style_string, "thin"))
    {
        return xlnt::border_style::thin;
    }
    else
    {
        std::string message = "unknown border style: ";
        message.append(border_style_string);
        throw std::runtime_error(border_style_string);
    }
}

} // namespace

namespace xlnt {

style_serializer::style_serializer(workbook &wb) : workbook_(wb)
{
}

protection style_serializer::read_protection(const xml_node &protection_node)
{
    protection prot;

    prot.set_locked(protection_type_from_string(protection_node.get_attribute("locked")));
    prot.set_hidden(protection_type_from_string(protection_node.get_attribute("hidden")));

    return prot;
}

alignment style_serializer::read_alignment(const xml_node &alignment_node)
{
    alignment align;

    align.set_wrap_text(is_true(alignment_node.get_attribute("wrapText")));
    align.set_shrink_to_fit(is_true(alignment_node.get_attribute("shrinkToFit")));

    bool has_vertical = alignment_node.has_attribute("vertical");

    if (has_vertical)
    {
        std::string vertical = alignment_node.get_attribute("vertical");

        if (vertical == "bottom")
        {
            align.set_vertical(vertical_alignment::bottom);
        }
        else if (vertical == "center")
        {
            align.set_vertical(vertical_alignment::center);
        }
        else if (vertical == "justify")
        {
            align.set_vertical(vertical_alignment::justify);
        }
        else if (vertical == "top")
        {
            align.set_vertical(vertical_alignment::top);
        }
        else
        {
            throw "unknown alignment";
        }
    }

    bool has_horizontal = alignment_node.has_attribute("horizontal");

    if (has_horizontal)
    {
        std::string horizontal = alignment_node.get_attribute("horizontal");

        if (horizontal == "left")
        {
            align.set_horizontal(horizontal_alignment::left);
        }
        else if (horizontal == "center")
        {
            align.set_horizontal(horizontal_alignment::center);
        }
        else if (horizontal == "center-continuous")
        {
            align.set_horizontal(horizontal_alignment::center_continuous);
        }
        else if (horizontal == "right")
        {
            align.set_horizontal(horizontal_alignment::right);
        }
        else if (horizontal == "justify")
        {
            align.set_horizontal(horizontal_alignment::justify);
        }
        else if (horizontal == "general")
        {
            align.set_horizontal(horizontal_alignment::general);
        }
        else
        {
            throw "unknown alignment";
        }
    }

    return align;
}

cell_style style_serializer::read_cell_style(const xml_node &style_node)
{
    cell_style s;
    
    // Alignment
    s.get_alignment().apply(style_node.has_child("alignment") || is_true(style_node.get_attribute("applyAlignment")));

    if (s.get_alignment().apply())
    {
        auto inline_alignment = read_alignment(style_node.get_child("alignment"));
        s.set_alignment(inline_alignment);
    }
    
    // Border
    s.get_border().apply(is_true(style_node.get_attribute("applyBorder")));
    auto border_index = style_node.has_attribute("borderId") ? std::stoull(style_node.get_attribute("borderId")) : 0;
    s.set_border(borders_.at(border_index));
    
    // Fill
    s.get_fill().apply(is_true(style_node.get_attribute("applyFill")));
    auto fill_index = style_node.has_attribute("fillId") ? std::stoull(style_node.get_attribute("fillId")) : 0;
    s.set_fill(fills_.at(fill_index));
    
    // Font
    s.get_font().apply(is_true(style_node.get_attribute("applyFont")));
    auto font_index = style_node.has_attribute("fontId") ? std::stoull(style_node.get_attribute("fontId")) : 0;
    s.set_font(fonts_.at(font_index));

    // Number Format
    s.get_number_format().apply(is_true(style_node.get_attribute("applyNumberFormat")));
    auto number_format_id = std::stoull(style_node.get_attribute("numFmtId"));

    bool builtin_format = true;

    for (const auto &num_fmt : number_formats_)
    {
        if (static_cast<std::size_t>(num_fmt.get_id()) == number_format_id)
        {
            s.set_number_format(num_fmt);
            builtin_format = false;
            break;
        }
    }

    if (builtin_format)
    {
        s.set_number_format(number_format::from_builtin_id(number_format_id));
    }

    // Protection
    s.get_protection().apply(style_node.has_attribute("protection") || is_true(style_node.get_attribute("applyProtection")));

    if (s.get_protection().apply())
    {
        auto inline_protection = read_protection(style_node.get_child("protection"));
        s.set_protection(inline_protection);
    }

    return s;
}

named_style style_serializer::read_named_style(const xml_node &named_style_node, const xml_node &style_parent_node)
{
    named_style s;

    s.set_name(named_style_node.get_attribute("name"));
    s.set_hidden(named_style_node.has_attribute("hidden") && is_true(named_style_node.get_attribute("hidden")));
    s.set_builtin_id(std::stoull(named_style_node.get_attribute("builtinId")));
    
    auto base_style_id = std::stoull(named_style_node.get_attribute("xfId"));
    auto base_style_node = style_parent_node.get_children().at(base_style_id);
    auto base_style = read_cell_style(base_style_node);
    
    //TODO shouldn't have to set apply after set_X()
    s.set_alignment(base_style.get_alignment());
    s.get_alignment().apply(base_style.get_alignment().apply());
    s.set_border(base_style.get_border());
    s.get_border().apply(base_style.get_border().apply());
    s.set_fill(base_style.get_fill());
    s.get_fill().apply(base_style.get_fill().apply());
    s.set_font(base_style.get_font());
    s.get_font().apply(base_style.get_font().apply());
    s.set_number_format(base_style.get_number_format());
    s.get_number_format().apply(base_style.get_number_format().apply());
    s.set_protection(base_style.get_protection());
    s.get_protection().apply(base_style.get_protection().apply());
    
    return s;
}

bool style_serializer::read_stylesheet(const xml_document &xml)
{
    auto stylesheet_node = xml.get_child("styleSheet");

    read_borders(stylesheet_node.get_child("borders"));
    read_fills(stylesheet_node.get_child("fills"));
    read_fonts(stylesheet_node.get_child("fonts"));
    read_number_formats(stylesheet_node.get_child("numFmts"));
    read_colors(stylesheet_node.get_child("colors"));
    read_named_styles(stylesheet_node.get_child("cellStyles"), stylesheet_node.get_child("cellStyleXfs"));
    read_cell_styles(stylesheet_node.get_child("cellXfs"));

    for (const auto &ns : named_styles_)
    {
        workbook_.create_named_style(ns.second.get_name()) = ns.second;
    }
    
    for (const auto &s : cell_styles_)
    {
        workbook_.add_style(s);
    }

    return true;
}

bool style_serializer::read_cell_styles(const xlnt::xml_node &cell_styles_node)
{
    for (auto style_node : cell_styles_node.get_children())
    {
        if (style_node.get_name() != "xf")
        {
            continue;
        }

        auto style = read_cell_style(style_node);
        
        if(style_node.has_attribute("xfId"))
        {
            auto named_style_index = std::stoull(style_node.get_attribute("xfId"));
            auto named_style = named_styles_.at(named_style_index);
            
            style.set_named_style(named_style.get_name());
        }
        
        cell_styles_.push_back(style);
    }
    
    return true;
}

bool style_serializer::read_named_styles(const xlnt::xml_node &named_styles_node, const xlnt::xml_node &cell_styles_node)
{
    for (auto named_style_node : named_styles_node.get_children())
    {
        auto ns = read_named_style(named_style_node, cell_styles_node);
        auto named_style_index = std::stoull(named_style_node.get_attribute("xfId"));
        
        named_styles_[named_style_index] = ns;
    }
    
    return true;
}

bool style_serializer::read_number_formats(const xml_node &number_formats_node)
{
    for (auto num_fmt_node : number_formats_node.get_children())
    {
        if (num_fmt_node.get_name() != "numFmt")
        {
            continue;
        }

        auto format_string = num_fmt_node.get_attribute("formatCode");

        if (format_string == "GENERAL")
        {
            format_string = "General";
        }

        number_format nf;

        nf.set_format_string(format_string);
        nf.set_id(std::stoull(num_fmt_node.get_attribute("numFmtId")));

        number_formats_.push_back(nf);
    }

    return true;
}

bool style_serializer::read_fonts(const xml_node &fonts_node)
{
    for (auto font_node : fonts_node.get_children())
    {
        fonts_.push_back(read_font(font_node));
    }

    return true;
}

font style_serializer::read_font(const xlnt::xml_node &font_node)
{
    font new_font;

    new_font.set_size(std::stoull(font_node.get_child("sz").get_attribute("val")));
    new_font.set_name(font_node.get_child("name").get_attribute("val"));

    if (font_node.has_child("color"))
    {
        new_font.set_color(read_color(font_node.get_child("color")));
    }

    if (font_node.has_child("family"))
    {
        new_font.set_family(std::stoull(font_node.get_child("family").get_attribute("val")));
    }

    if (font_node.has_child("scheme"))
    {
        new_font.set_scheme(font_node.get_child("scheme").get_attribute("val"));
    }

    if (font_node.has_child("b"))
    {
        if(font_node.get_child("b").has_attribute("val"))
        {
            new_font.set_bold(is_true(font_node.get_child("b").get_attribute("val")));
        }
        else
        {
            new_font.set_bold(true);
        }
    }

    if (font_node.has_child("strike"))
    {
        if(font_node.get_child("strike").has_attribute("val"))
        {
            new_font.set_strikethrough(is_true(font_node.get_child("strike").get_attribute("val")));
        }
        else
        {
            new_font.set_strikethrough(true);
        }
    }

    if (font_node.has_child("i"))
    {
        if(font_node.get_child("i").has_attribute("val"))
        {
            new_font.set_italic(is_true(font_node.get_child("i").get_attribute("val")));
        }
        else
        {
            new_font.set_italic(true);
        }
    }

    if (font_node.has_child("u"))
    {
        if (font_node.get_child("u").has_attribute("val"))
        {
            std::string underline_string = font_node.get_child("u").get_attribute("val");
            new_font.set_underline(underline_style_from_string(underline_string));
        }
        else
        {
            new_font.set_underline(xlnt::font::underline_style::single);
        }
    }

    return new_font;
}

bool style_serializer::read_colors(const xml_node &colors_node)
{
    if (colors_node.has_child("indexedColors"))
    {
        auto indexed_colors = read_indexed_colors(colors_node.get_child("indexedColors"));

        for (const auto &indexed_color : indexed_colors)
        {
            colors_.push_back(indexed_color);
        }
    }

    return true;
}

bool style_serializer::read_fills(const xml_node &fills_node)
{
    for (auto fill_node : fills_node.get_children())
    {
        fills_.push_back(read_fill(fill_node));
    }

    return true;
}

fill style_serializer::read_fill(const xml_node &fill_node)
{
    fill new_fill;

    if (fill_node.has_child("patternFill"))
    {
        new_fill.set_type(fill::type::pattern);

        auto pattern_fill_node = fill_node.get_child("patternFill");
        new_fill.set_pattern_type(pattern_fill_type_from_string(pattern_fill_node.get_attribute("patternType")));

        if (pattern_fill_node.has_child("bgColor"))
        {
            new_fill.get_background_color() = read_color(pattern_fill_node.get_child("bgColor"));
        }

        if (pattern_fill_node.has_child("fgColor"))
        {
            new_fill.get_foreground_color() = read_color(pattern_fill_node.get_child("fgColor"));
        }
    }

    return new_fill;
}

side style_serializer::read_side(const xml_node &side_node)
{
    side new_side;

    if (side_node.has_attribute("style"))
    {
        new_side.get_border_style() = border_style_from_string(side_node.get_attribute("style"));
    }

    if (side_node.has_child("color"))
    {
        new_side.get_color() = read_color(side_node.get_child("color"));
    }

    return new_side;
}

bool style_serializer::read_borders(const xml_node &borders_node)
{
    for (auto border_node : borders_node.get_children())
    {
        borders_.push_back(read_border(border_node));
    }

    return true;
}

border style_serializer::read_border(const xml_node &border_node)
{
    border new_border;

    if (border_node.has_child("start"))
    {
        new_border.get_start() = read_side(border_node.get_child("start"));
    }

    if (border_node.has_child("end"))
    {
        new_border.get_end() = read_side(border_node.get_child("end"));
    }

    if (border_node.has_child("left"))
    {
        new_border.get_left() = read_side(border_node.get_child("left"));
    }

    if (border_node.has_child("right"))
    {
        new_border.get_right() = read_side(border_node.get_child("right"));
    }

    if (border_node.has_child("top"))
    {
        new_border.get_top() = read_side(border_node.get_child("top"));
    }

    if (border_node.has_child("bottom"))
    {
        new_border.get_bottom() = read_side(border_node.get_child("bottom"));
    }

    if (border_node.has_child("diagonal"))
    {
        new_border.get_diagonal() = read_side(border_node.get_child("diagonal"));
    }

    if (border_node.has_child("vertical"))
    {
        new_border.get_vertical() = read_side(border_node.get_child("vertical"));
    }

    if (border_node.has_child("horizontal"))
    {
        new_border.get_horizontal() = read_side(border_node.get_child("horizontal"));
    }

    return new_border;
}

std::vector<color> style_serializer::read_indexed_colors(const xml_node &indexed_colors_node)
{
    std::vector<color> colors;

    for (auto color_node : indexed_colors_node.get_children())
    {
        colors.push_back(read_color(color_node));
    }

    return colors;
}

color style_serializer::read_color(const xml_node &color_node)
{
    if (color_node.has_attribute("rgb"))
    {
        return color(color::type::rgb, color_node.get_attribute("rgb"));
    }
    else if (color_node.has_attribute("theme"))
    {
        return color(color::type::theme, std::stoull(color_node.get_attribute("theme")));
    }
    else if (color_node.has_attribute("indexed"))
    {
        return color(color::type::indexed, std::stoull(color_node.get_attribute("indexed")));
    }
    else if (color_node.has_attribute("auto"))
    {
        return color(color::type::auto_, std::stoull(color_node.get_attribute("auto")));
    }

    throw std::runtime_error("bad color");
}

void style_serializer::initialize_vectors()
{
    std::unordered_set<cell_style, std::hash<hashable>> cell_styles_set;
    
    for (auto ws : workbook_)
    {
        for (auto row : ws)
        {
            for (auto c : row)
            {
                if (c.has_style())
                {
                    cell_styles_set.insert(c.get_style());
                }
            }
        }
    }
    
    cell_styles_.assign(cell_styles_set.begin(), cell_styles_set.end());
    colors_.clear();
    borders_.clear();
    fills_.clear();
    fonts_.clear();
    number_formats_.clear();
    
    for (auto &style : cell_styles_)
    {
        if (std::find(borders_.begin(), borders_.end(), style.get_border()) == borders_.end())
        {
            borders_.push_back(style.get_border());
        }
        
        if (std::find(fills_.begin(), fills_.end(), style.get_fill()) == fills_.end())
        {
            fills_.push_back(style.get_fill());
        }
        
        if (std::find(fonts_.begin(), fonts_.end(), style.get_font()) == fonts_.end())
        {
            fonts_.push_back(style.get_font());
        }
        
        if (std::find(number_formats_.begin(), number_formats_.end(), style.get_number_format()) == number_formats_.end())
        {
            number_formats_.push_back(style.get_number_format());
        }
    }
    
    if (fonts_.empty())
    {
        fonts_.push_back(font());
    }
}

bool style_serializer::write_fonts(xlnt::xml_node &fonts_node) const
{
    fonts_node.add_attribute("count", std::to_string(fonts_.size()));
    // TODO: what does this do?
    // fonts_node.add_attribute("x14ac:knownFonts", "1");

    for (auto &f : fonts_)
    {
        auto font_node = fonts_node.add_child("font");

        if (f.is_bold())
        {
            auto bold_node = font_node.add_child("b");
            bold_node.add_attribute("val", "1");
        }

        if (f.is_italic())
        {
            auto italic_node = font_node.add_child("i");
            italic_node.add_attribute("val", "1");
        }

        if (f.is_underline())
        {
            auto underline_node = font_node.add_child("u");

            switch (f.get_underline())
            {
            case font::underline_style::single:
                underline_node.add_attribute("val", "single");
                break;
            default:
                break;
            }
        }

        if (f.is_strikethrough())
        {
            auto strike_node = font_node.add_child("strike");
            strike_node.add_attribute("val", "1");
        }

        auto size_node = font_node.add_child("sz");
        size_node.add_attribute("val", std::to_string(f.get_size()));

        auto color_node = font_node.add_child("color");

        if (f.get_color().get_type() == color::type::indexed)
        {
            color_node.add_attribute("indexed", std::to_string(f.get_color().get_index()));
        }
        else if (f.get_color().get_type() == color::type::theme)
        {
            color_node.add_attribute("theme", std::to_string(f.get_color().get_theme()));
        }
        else if (f.get_color().get_type() == color::type::auto_)
        {
            color_node.add_attribute("auto", std::to_string(f.get_color().get_auto()));
        }
        else if (f.get_color().get_type() == color::type::rgb)
        {
            color_node.add_attribute("rgb", f.get_color().get_rgb_string());
        }
        
        auto name_node = font_node.add_child("name");
        name_node.add_attribute("val", f.get_name());

        if (f.has_family())
        {
            auto family_node = font_node.add_child("family");
            family_node.add_attribute("val", std::to_string(f.get_family()));
        }

        if (f.has_scheme())
        {
            auto scheme_node = font_node.add_child("scheme");
            scheme_node.add_attribute("val", "minor");
        }
    }
    
    return true;
}

bool style_serializer::write_fills(xlnt::xml_node &fills_node) const
{
    fills_node.add_attribute("count", std::to_string(fills_.size()));

    for (auto &fill_ : fills_)
    {
        auto fill_node = fills_node.add_child("fill");

        if (fill_.get_type() == fill::type::pattern)
        {
            auto pattern_fill_node = fill_node.add_child("patternFill");
            auto type_string = fill_.get_pattern_type_string();
            pattern_fill_node.add_attribute("patternType", type_string);
            
            if (fill_.get_pattern_type() != xlnt::fill::pattern_type::solid) continue;

            if (fill_.get_foreground_color())
            {
                auto fg_color_node = pattern_fill_node.add_child("fgColor");

                switch (fill_.get_foreground_color()->get_type())
                {
                case color::type::auto_:
                    fg_color_node.add_attribute("auto", std::to_string(fill_.get_foreground_color()->get_auto()));
                    break;
                case color::type::theme:
                    fg_color_node.add_attribute("theme", std::to_string(fill_.get_foreground_color()->get_theme()));
                    break;
                case color::type::indexed:
                    fg_color_node.add_attribute("indexed", std::to_string(fill_.get_foreground_color()->get_index()));
                    break;
                case color::type::rgb:
                    fg_color_node.add_attribute("rgb", fill_.get_foreground_color()->get_rgb_string());
                    break;
                default:
                    throw std::runtime_error("bad type");
                }
            }

            if (fill_.get_background_color())
            {
                auto bg_color_node = pattern_fill_node.add_child("bgColor");

                switch (fill_.get_background_color()->get_type())
                {
                case color::type::auto_:
                    bg_color_node.add_attribute("auto", std::to_string(fill_.get_background_color()->get_auto()));
                    break;
                case color::type::theme:
                    bg_color_node.add_attribute("theme", std::to_string(fill_.get_background_color()->get_theme()));
                    break;
                case color::type::indexed:
                    bg_color_node.add_attribute("indexed", std::to_string(fill_.get_background_color()->get_index()));
                    break;
                case color::type::rgb:
                    bg_color_node.add_attribute("rgb", fill_.get_background_color()->get_rgb_string());
                    break;
                default:
                    throw std::runtime_error("bad type");
                }
            }
        }
        else if (fill_.get_type() == fill::type::solid)
        {
            auto solid_fill_node = fill_node.add_child("solidFill");
            solid_fill_node.add_child("color");
        }
        else if (fill_.get_type() == fill::type::gradient)
        {
            auto gradient_fill_node = fill_node.add_child("gradientFill");

            if (fill_.get_gradient_type_string() == "linear")
            {
                gradient_fill_node.add_attribute("degree", std::to_string(fill_.get_rotation()));
            }
            else if (fill_.get_gradient_type_string() == "path")
            {
                gradient_fill_node.add_attribute("left", std::to_string(fill_.get_gradient_left()));
                gradient_fill_node.add_attribute("right", std::to_string(fill_.get_gradient_right()));
                gradient_fill_node.add_attribute("top", std::to_string(fill_.get_gradient_top()));
                gradient_fill_node.add_attribute("bottom", std::to_string(fill_.get_gradient_bottom()));

                auto start_node = gradient_fill_node.add_child("stop");
                start_node.add_attribute("position", "0");

                auto end_node = gradient_fill_node.add_child("stop");
                end_node.add_attribute("position", "1");
            }
        }
    }

    return true;
}

bool style_serializer::write_borders(xlnt::xml_node &borders_node) const
{
    borders_node.add_attribute("count", std::to_string(borders_.size()));

    for (const auto &border_ : borders_)
    {
        auto border_node = borders_node.add_child("border");

        std::vector<std::tuple<std::string, const std::experimental::optional<side>>> sides;
        
        sides.push_back(std::make_tuple("start", border_.get_start()));
        sides.push_back(std::make_tuple("end", border_.get_end()));
        sides.push_back(std::make_tuple("left", border_.get_left()));
        sides.push_back(std::make_tuple("right", border_.get_right()));
        sides.push_back(std::make_tuple("top", border_.get_top()));
        sides.push_back(std::make_tuple("bottom", border_.get_bottom()));
        sides.push_back(std::make_tuple("diagonal", border_.get_diagonal()));
        sides.push_back(std::make_tuple("vertical", border_.get_vertical()));
        sides.push_back(std::make_tuple("horizontal", border_.get_horizontal()));

        for (const auto &side_tuple : sides)
        {
            std::string current_name = std::get<0>(side_tuple);
            const auto current_side = std::get<1>(side_tuple);

            if (current_side)
            {
                auto side_node = border_node.add_child(current_name);

                if (current_side->get_border_style())
                {
                    std::string style_string;

                    switch (*current_side->get_border_style())
                    {
                    case border_style::none:
                        style_string = "none";
                        break;
                    case border_style::dashdot:
                        style_string = "dashDot";
                        break;
                    case border_style::dashdotdot:
                        style_string = "dashDotDot";
                        break;
                    case border_style::dashed:
                        style_string = "dashed";
                        break;
                    case border_style::dotted:
                        style_string = "dotted";
                        break;
                    case border_style::double_:
                        style_string = "double";
                        break;
                    case border_style::hair:
                        style_string = "hair";
                        break;
                    case border_style::medium:
                        style_string = "thin";
                        break;
                    case border_style::mediumdashdot:
                        style_string = "mediumDashDot";
                        break;
                    case border_style::mediumdashdotdot:
                        style_string = "mediumDashDotDot";
                        break;
                    case border_style::mediumdashed:
                        style_string = "mediumDashed";
                        break;
                    case border_style::slantdashdot:
                        style_string = "slantDashDot";
                        break;
                    case border_style::thick:
                        style_string = "thick";
                        break;
                    case border_style::thin:
                        style_string = "thin";
                        break;
                    default:
                        throw std::runtime_error("invalid style");
                    }

                    side_node.add_attribute("style", style_string);
                }

                if (current_side->get_color())
                {
                    auto color_node = side_node.add_child("color");

                    if (current_side->get_color()->get_type() == color::type::indexed)
                    {
                        color_node.add_attribute("indexed", std::to_string(current_side->get_color()->get_index()));
                    }
                    else if (current_side->get_color()->get_type() == color::type::theme)
                    {
                        color_node.add_attribute("theme", std::to_string(current_side->get_color()->get_theme()));
                    }
                    else if (current_side->get_color()->get_type() == color::type::rgb)
                    {
                        color_node.add_attribute("rgb", current_side->get_color()->get_rgb_string());
                    }
                    else
                    {
                        throw std::runtime_error("invalid color type");
                    }
                }
            }
        }
    }

    return true;
}

bool style_serializer::write_style_common(const alignment &style_alignment, const border &style_border, const fill &style_fill, const font &style_font, const number_format &style_number_format, const protection &style_protection, xml_node xf_node) const
{
    xf_node.add_attribute("numFmtId", std::to_string(style_number_format.get_id()));
     
    auto font_id = std::distance(fonts_.begin(), std::find(fonts_.begin(), fonts_.end(), style_font));
    xf_node.add_attribute("fontId", std::to_string(font_id));

    if (style_fill.apply())
    {
        auto fill_id = std::distance(fills_.begin(), std::find(fills_.begin(), fills_.end(), style_fill));
        xf_node.add_attribute("fillId", std::to_string(fill_id));
    }

    if (style_border.apply())
    {
        auto border_id = std::distance(borders_.begin(), std::find(borders_.begin(), borders_.end(), style_border));
        xf_node.add_attribute("borderId", std::to_string(border_id));
    }

    xf_node.add_attribute("applyNumberFormat", style_number_format.apply() ? "1" : "0");
    xf_node.add_attribute("applyFill", style_fill.apply() ? "1" : "0");
    xf_node.add_attribute("applyFont", style_font.apply() ? "1" : "0");
    xf_node.add_attribute("applyBorder", style_border.apply() ? "1" : "0");
    xf_node.add_attribute("applyAlignment", style_alignment.apply() ? "1" : "0");
    xf_node.add_attribute("applyProtection", style_protection.apply() ? "1" : "0");

    if (style_alignment.apply())
    {
        auto alignment_node = xf_node.add_child("alignment");

        if (style_alignment.has_vertical())
        {
            switch (style_alignment.get_vertical())
            {
            case vertical_alignment::bottom:
                alignment_node.add_attribute("vertical", "bottom");
                break;
            case vertical_alignment::center:
                alignment_node.add_attribute("vertical", "center");
                break;
            case vertical_alignment::justify:
                alignment_node.add_attribute("vertical", "justify");
                break;
            case vertical_alignment::top:
                alignment_node.add_attribute("vertical", "top");
                break;
            default:
                throw std::runtime_error("invalid alignment");
            }
        }

        if (style_alignment.has_horizontal())
        {
            switch (style_alignment.get_horizontal())
            {
            case horizontal_alignment::center:
                alignment_node.add_attribute("horizontal", "center");
                break;
            case horizontal_alignment::center_continuous:
                alignment_node.add_attribute("horizontal", "center_continuous");
                break;
            case horizontal_alignment::general:
                alignment_node.add_attribute("horizontal", "general");
                break;
            case horizontal_alignment::justify:
                alignment_node.add_attribute("horizontal", "justify");
                break;
            case horizontal_alignment::left:
                alignment_node.add_attribute("horizontal", "left");
                break;
            case horizontal_alignment::right:
                alignment_node.add_attribute("horizontal", "right");
                break;
            default:
                throw std::runtime_error("invalid alignment");
            }
        }

        if (style_alignment.get_wrap_text())
        {
            alignment_node.add_attribute("wrapText", "1");
        }
        
        if (style_alignment.get_shrink_to_fit())
        {
            alignment_node.add_attribute("shrinkToFit", "1");
        }
    }
    
    return true;
}

bool style_serializer::write_cell_style(const cell_style &style, xml_node &xf_node) const
{
    return write_style_common(style.get_alignment(), style.get_border(), style.get_fill(), style.get_font(), style.get_number_format(), style.get_protection(), xf_node);
}

bool style_serializer::write_named_style(const named_style &style, xml_node &xf_node) const
{
    return write_style_common(style.get_alignment(), style.get_border(), style.get_fill(), style.get_font(), style.get_number_format(), style.get_protection(), xf_node);
}

bool style_serializer::write_named_styles(xml_node &cell_styles_node, xml_node &styles_node) const
{
    styles_node.add_attribute("count", std::to_string(named_styles_.size()));

    for(auto &key_style : named_styles_)
    {
        auto xf_node = styles_node.add_child("xf");
        write_named_style(key_style.second, xf_node);
    }
    
    cell_styles_node.add_attribute("count", std::to_string(named_styles_.size()));

    for(auto &key_style : named_styles_)
    {
        auto cell_style_node = cell_styles_node.add_child("cellStyle");
        
        cell_style_node.add_attribute("name", key_style.second.get_name());
        cell_style_node.add_attribute("xfId", std::to_string(key_style.first));
        cell_style_node.add_attribute("builtinId", std::to_string(key_style.second.get_builtin_id()));
        
        if (key_style.second.get_hidden())
        {
            cell_style_node.add_attribute("hidden", "1");
        }
    }

    return true;
}

bool style_serializer::write_cell_styles(xml_node &cell_styles_node) const
{
    cell_styles_node.add_attribute("count", std::to_string(cell_styles_.size()));

    for(auto &style : cell_styles_)
    {
        auto xf_node = cell_styles_node.add_child("xf");
        write_cell_style(style, xf_node);
        xf_node.add_attribute("xfId", "0"); //TODO point to named style here
    }

    return true;
}

bool write_dxfs(xlnt::xml_node &dxfs_node)
{
    dxfs_node.add_attribute("count", "0");
    return true;
}

bool write_table_styles(xlnt::xml_node &table_styles_node)
{
    table_styles_node.add_attribute("count", "0");
    table_styles_node.add_attribute("defaultTableStyle", "TableStyleMedium2");
    table_styles_node.add_attribute("defaultPivotStyle", "PivotStyleMedium9");
    
    return true;
}

bool style_serializer::write_colors(xlnt::xml_node &colors_node) const
{
    auto indexed_colors_node = colors_node.add_child("indexedColors");

    for (auto &c : colors_)
    {
        indexed_colors_node.add_child("rgbColor").add_attribute("rgb", c.get_rgb_string());
    }
    
    return true;
}

bool style_serializer::write_ext_list(xlnt::xml_node &ext_list_node) const
{
    auto ext_node = ext_list_node.add_child("ext");
    ext_node.add_attribute("uri", "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}");
    ext_node.add_attribute("xmlns:x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
    ext_node.add_child("x14:slicerStyles").add_attribute("defaultSlicerStyle", "SlicerStyleLight1");
    
    return true;
}

xml_document style_serializer::write_stylesheet()
{
    initialize_vectors();
    
    xml_document doc;
    auto root_node = doc.add_child("styleSheet");
    doc.add_namespace("", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    doc.add_namespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
    root_node.add_attribute("mc:Ignorable", "x14ac");
    doc.add_namespace("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

    auto number_formats_node = root_node.add_child("numFmts");
    write_number_formats(number_formats_node);

    auto fonts_node = root_node.add_child("fonts");
    write_fonts(fonts_node);

    auto fills_node = root_node.add_child("fills");
    write_fills(fills_node);
    
    auto borders_node = root_node.add_child("borders");
    write_borders(borders_node);
    
    auto cell_style_xfs_node = root_node.add_child("cellStyleXfs");
    auto cell_styles_node = root_node.add_child("cellStyles");
    write_named_styles(root_node, cell_style_xfs_node);

    auto cell_xfs_node = root_node.add_child("cellXfs");
    write_cell_styles(cell_styles_node);
    
    auto dxfs_node = root_node.add_child("dxfs");
    write_dxfs(dxfs_node);

    auto table_styles_node = root_node.add_child("tableStyles");
    write_table_styles(table_styles_node);

    if(!colors_.empty())
    {
        auto colors_node = root_node.add_child("colors");
        write_colors(colors_node);
    }
    
    auto ext_list_node = root_node.add_child("extLst");
    write_ext_list(ext_list_node);

    return doc;
}

bool style_serializer::write_number_formats(xml_node &number_formats_node) const
{
    number_formats_node.add_attribute("count", std::to_string(number_formats_.size()));

    for (const auto &num_fmt : number_formats_)
    {
        auto num_fmt_node = number_formats_node.add_child("numFmt");
        num_fmt_node.add_attribute("numFmtId", std::to_string(num_fmt.get_id()));
        num_fmt_node.add_attribute("formatCode", num_fmt.get_format_string());
    }

    return true;
}

const std::vector<border> &style_serializer::get_borders() const
{
    return borders_;
}

const std::vector<fill> &style_serializer::get_fills() const
{
    return fills_;
}

const std::vector<font> &style_serializer::get_fonts() const
{
    return fonts_;
}

const std::vector<number_format> &style_serializer::get_number_formats() const
{
    return number_formats_;
}

const std::vector<color> &style_serializer::get_colors() const
{
    return colors_;
}

} // namespace xlnt
