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
#include <xlnt/serialization/style_serializer.hpp>
#include <xlnt/serialization/xml_document.hpp>
#include <xlnt/serialization/xml_node.hpp>
#include <xlnt/styles/alignment.hpp>
#include <xlnt/styles/border.hpp>
#include <xlnt/styles/fill.hpp>
#include <xlnt/styles/font.hpp>
#include <xlnt/styles/number_format.hpp>
#include <xlnt/styles/protection.hpp>
#include <xlnt/styles/style.hpp>

namespace {

bool is_true(const std::string &bool_string)
{
    return bool_string == "1";
}

xlnt::protection::type protection_type_from_string(const std::string &type_string)
{
    if (type_string == "true")
    {
        return xlnt::protection::type::protected_;
    }
    else if (type_string == "inherit")
    {
        return xlnt::protection::type::inherit;
    }

    return xlnt::protection::type::unprotected;
};

xlnt::font::underline_style underline_style_from_string(const std::string &underline_string)
{
    if (underline_string == "none")
    {
        return xlnt::font::underline_style::none;
    }
    else if (underline_string == "single")
    {
        return xlnt::font::underline_style::single;
    }
    else if (underline_string == "single-accounting")
    {
        return xlnt::font::underline_style::single_accounting;
    }
    else if (underline_string == "double")
    {
        return xlnt::font::underline_style::double_;
    }
    else if (underline_string == "double-accounting")
    {
        return xlnt::font::underline_style::double_accounting;
    }

    return xlnt::font::underline_style::none;
}

xlnt::fill::pattern_type pattern_fill_type_from_string(const std::string &fill_type)
{
    if (fill_type == "none") return xlnt::fill::pattern_type::none;
    if (fill_type == "solid") return xlnt::fill::pattern_type::solid;
    if (fill_type == "gray125") return xlnt::fill::pattern_type::gray125;
    return xlnt::fill::pattern_type::none;
};

xlnt::border_style border_style_from_string(const std::string &border_style_string)
{
    if (border_style_string == "none")
    {
        return xlnt::border_style::none;
    }
    else if (border_style_string == "dashdot")
    {
        return xlnt::border_style::dashdot;
    }
    else if (border_style_string == "dashdotdot")
    {
        return xlnt::border_style::dashdotdot;
    }
    else if (border_style_string == "dashed")
    {
        return xlnt::border_style::dashed;
    }
    else if (border_style_string == "dotted")
    {
        return xlnt::border_style::dotted;
    }
    else if (border_style_string == "double")
    {
        return xlnt::border_style::double_;
    }
    else if (border_style_string == "hair")
    {
        return xlnt::border_style::hair;
    }
    else if (border_style_string == "medium")
    {
        return xlnt::border_style::medium;
    }
    else if (border_style_string == "mediumdashdot")
    {
        return xlnt::border_style::mediumdashdot;
    }
    else if (border_style_string == "mediumdashdotdot")
    {
        return xlnt::border_style::mediumdashdotdot;
    }
    else if (border_style_string == "mediumdashed")
    {
        return xlnt::border_style::mediumdashed;
    }
    else if (border_style_string == "slantdashdot")
    {
        return xlnt::border_style::slantdashdot;
    }
    else if (border_style_string == "thick")
    {
        return xlnt::border_style::thick;
    }
    else if (border_style_string == "thin")
    {
        return xlnt::border_style::thin;
    }
    else
    {
        throw std::runtime_error("unknown border style");
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

style style_serializer::read_style(const xml_node &style_node)
{
    style s;

    s.apply_number_format(is_true(style_node.get_attribute("applyNumberFormat")));
    s.number_format_id_ = std::stoull(style_node.get_attribute("numFmtId"));

    bool builtin_format = true;

    for (auto num_fmt : workbook_.get_number_formats())
    {
        if (static_cast<std::size_t>(num_fmt.get_id()) == s.get_number_format_id())
        {
            s.number_format_ = num_fmt;
            builtin_format = false;
            break;
        }
    }

    if (builtin_format)
    {
        s.number_format_ = number_format::from_builtin_id(s.get_number_format_id());
    }

    s.apply_font(is_true(style_node.get_attribute("applyFont")));
    s.font_id_ = style_node.has_attribute("fontId") ? std::stoull(style_node.get_attribute("fontId")) : 0;
    s.font_ = workbook_.get_font(s.font_id_);

    s.apply_fill(is_true(style_node.get_attribute("applyFill")));
    s.fill_id_ = style_node.has_attribute("fillId") ? std::stoull(style_node.get_attribute("fillId")) : 0;
    s.fill_ = workbook_.get_fill(s.fill_id_);

    s.apply_border(is_true(style_node.get_attribute("applyBorder")));
    s.border_id_ = style_node.has_attribute("borderId") ? std::stoull(style_node.get_attribute("borderId")) : 0;
    s.border_ = workbook_.get_border(s.border_id_);

    s.apply_protection(style_node.has_attribute("protection"));

    if (s.protection_apply_)
    {
        auto inline_protection = read_protection(style_node.get_child("protection"));
        s.protection_ = inline_protection;
    }

    s.apply_alignment(style_node.has_child("alignment"));

    if (s.alignment_apply_)
    {
        auto inline_alignment = read_alignment(style_node.get_child("alignment"));
        s.alignment_ = inline_alignment;
    }

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

    auto cell_xfs_node = stylesheet_node.get_child("cellXfs");

    for (auto xf_node : cell_xfs_node.get_children())
    {
        if (xf_node.get_name() != "xf")
        {
            continue;
        }

        workbook_.add_style(read_style(xf_node));
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

        workbook_.add_number_format(nf);
    }

    return true;
}

bool style_serializer::read_fonts(const xml_node &fonts_node)
{
    for (auto font_node : fonts_node.get_children())
    {
        workbook_.add_font(read_font(font_node));
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
        new_font.set_bold(is_true(font_node.get_child("b").get_attribute("val")));
    }

    if (font_node.has_child("strike"))
    {
        new_font.set_strikethrough(is_true(font_node.get_child("strike").get_attribute("val")));
    }

    if (font_node.has_child("i"))
    {
        new_font.set_italic(is_true(font_node.get_child("i").get_attribute("val")));
    }

    if (font_node.has_child("u"))
    {
        std::string underline_string = font_node.get_child("u").get_attribute("val");
        new_font.set_underline(underline_style_from_string(underline_string));
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
            workbook_.add_color(indexed_color);
        }
    }

    return true;
}

bool style_serializer::read_fills(const xml_node &fills_node)
{
    for (auto fill_node : fills_node.get_children())
    {
        workbook_.add_fill(read_fill(fill_node));
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
        workbook_.add_border(read_border(border_node));
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

xml_document style_serializer::write_stylesheet() const
{
    xml_document xml;

    auto style_sheet_node = xml.add_child("styleSheet");
    xml.add_namespace("", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    xml.add_namespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
    style_sheet_node.add_attribute("mc:Ignorable", "x14ac");
    xml.add_namespace("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

    auto num_fmts_node = style_sheet_node.add_child("numFmts");
    auto num_fmts = workbook_.get_number_formats();
    num_fmts_node.add_attribute("count", std::to_string(num_fmts.size()));

    for (const auto &num_fmt : num_fmts)
    {
        auto num_fmt_node = num_fmts_node.add_child("numFmt");
        num_fmt_node.add_attribute("numFmtId", std::to_string(num_fmt.get_id()));
        num_fmt_node.add_attribute("formatCode", num_fmt.get_format_string());
    }

    auto fonts_node = style_sheet_node.add_child("fonts");
    auto fonts = workbook_.get_fonts();

    if (fonts.empty())
    {
        fonts.push_back(font());
    }

    fonts_node.add_attribute("count", std::to_string(fonts.size()));
    // TODO: what does this do?
    // fonts_node.add_attribute("x14ac:knownFonts", "1");

    for (auto &f : fonts)
    {
        auto font_node = fonts_node.add_child("font");

        if (f.is_bold())
        {
            auto bold_node = font_node.add_child("b");
            bold_node.add_attribute("val", "1");
        }

        if (f.is_italic())
        {
            auto bold_node = font_node.add_child("i");
            bold_node.add_attribute("val", "1");
        }

        if (f.is_underline())
        {
            auto bold_node = font_node.add_child("u");

            switch (f.get_underline())
            {
            case font::underline_style::single:
                bold_node.add_attribute("val", "single");
                break;
            default:
                break;
            }
        }

        if (f.is_strikethrough())
        {
            auto bold_node = font_node.add_child("strike");
            bold_node.add_attribute("val", "1");
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

    auto fills_node = style_sheet_node.add_child("fills");
    const auto &fills = workbook_.get_fills();
    fills_node.add_attribute("count", std::to_string(fills.size()));

    for (auto &fill_ : fills)
    {
        auto fill_node = fills_node.add_child("fill");

        if (fill_.get_type() == fill::type::pattern)
        {
            auto pattern_fill_node = fill_node.add_child("patternFill");
            auto type_string = fill_.get_pattern_type_string();
            pattern_fill_node.add_attribute("patternType", type_string);

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

    auto borders_node = style_sheet_node.add_child("borders");
    const auto &borders = workbook_.get_borders();
    borders_node.add_attribute("count", std::to_string(borders.size()));

    for (const auto &border_ : borders)
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
                        style_string = "dashdot";
                        break;
                    case border_style::dashdotdot:
                        style_string = "dashdotdot";
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
                        style_string = "mediumdashdot";
                        break;
                    case border_style::mediumdashdotdot:
                        style_string = "mediumdashdotdot";
                        break;
                    case border_style::mediumdashed:
                        style_string = "mediumdashed";
                        break;
                    case border_style::slantdashdot:
                        style_string = "slantdashdot";
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
                    else
                    {
                        throw std::runtime_error("invalid color type");
                    }
                }
            }
        }
    }

    auto cell_style_xfs_node = style_sheet_node.add_child("cellStyleXfs");
    cell_style_xfs_node.add_attribute("count", "1");

    auto style_xf_node = cell_style_xfs_node.add_child("xf");
    style_xf_node.add_attribute("numFmtId", "0");
    style_xf_node.add_attribute("fontId", "0");
    style_xf_node.add_attribute("fillId", "0");
    style_xf_node.add_attribute("borderId", "0");

    auto cell_xfs_node = style_sheet_node.add_child("cellXfs");
    const auto &styles = workbook_.get_styles();
    cell_xfs_node.add_attribute("count", std::to_string(styles.size()));

    for (auto &style : styles)
    {
        auto xf_node = cell_xfs_node.add_child("xf");
        xf_node.add_attribute("numFmtId", std::to_string(style.get_number_format().get_id()));
        xf_node.add_attribute("fontId", std::to_string(style.get_font_id()));

        if (style.fill_apply_)
        {
            xf_node.add_attribute("fillId", std::to_string(style.get_fill_id()));
        }

        if (style.border_apply_)
        {
            xf_node.add_attribute("borderId", std::to_string(style.get_border_id()));
        }

        xf_node.add_attribute("applyNumberFormat", style.number_format_apply_ ? "1" : "0");
        xf_node.add_attribute("applyFont", style.font_apply_ ? "1" : "0");
        xf_node.add_attribute("applyFill", style.fill_apply_ ? "1" : "0");
        xf_node.add_attribute("applyBorder", style.border_apply_ ? "1" : "0");
        xf_node.add_attribute("applyAlignment", style.alignment_apply_ ? "1" : "0");
        xf_node.add_attribute("applyProtection", style.protection_apply_ ? "1" : "0");

        if (style.alignment_apply_)
        {
            auto alignment_node = xf_node.add_child("alignment");

            if (style.alignment_.has_vertical())
            {
                switch (style.alignment_.get_vertical())
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

            if (style.alignment_.has_horizontal())
            {
                switch (style.alignment_.get_horizontal())
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

            if (style.alignment_.get_wrap_text())
            {
                alignment_node.add_attribute("wrapText", "1");
            }
        }
    }

    auto cell_styles_node = style_sheet_node.add_child("cellStyles");
    cell_styles_node.add_attribute("count", "1");
    auto cell_style_node = cell_styles_node.add_child("cellStyle");
    cell_style_node.add_attribute("name", "Normal");
    cell_style_node.add_attribute("xfId", "0");
    cell_style_node.add_attribute("builtinId", "0");

    style_sheet_node.add_child("dxfs").add_attribute("count", "0");

    auto table_styles_node = style_sheet_node.add_child("tableStyles");
    table_styles_node.add_attribute("count", "0");
    table_styles_node.add_attribute("defaultTableStyle", "TableStyleMedium2");
    table_styles_node.add_attribute("defaultPivotStyle", "PivotStyleMedium9");

    auto colors_node = style_sheet_node.add_child("colors");
    auto indexed_colors_node = colors_node.add_child("indexedColors");

    for (auto &c : workbook_.get_colors())
    {
        indexed_colors_node.add_child("rgbColor").add_attribute("rgb", c.get_rgb_string());
    }

    auto ext_list_node = style_sheet_node.add_child("extLst");
    auto ext_node = ext_list_node.add_child("ext");
    ext_node.add_attribute("uri", "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}");
    ext_node.add_attribute("xmlns:x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
    ext_node.add_child("x14:slicerStyles").add_attribute("defaultSlicerStyle", "SlicerStyleLight1");

    return xml;
}

bool style_serializer::write_number_formats(xml_node number_formats_node) const
{
    number_formats_node.add_attribute("count", std::to_string(workbook_.get_number_formats().size()));

    for (const auto &num_fmt : workbook_.get_number_formats())
    {
        auto num_fmt_node = number_formats_node.add_child("numFmt");
        num_fmt_node.add_attribute("numFmtId", std::to_string(num_fmt.get_id()));
        num_fmt_node.add_attribute("formatCode", num_fmt.get_format_string());
    }

    return true;
}

} // namespace xlnt
