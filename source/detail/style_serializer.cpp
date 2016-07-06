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
#include <cctype> // for std::tolower
#include <pugixml.hpp>
#include <unordered_set>
#include <utility> // for std::move

#include <detail/style_serializer.hpp>
#include <detail/stylesheet.hpp>
#include <detail/workbook_impl.hpp>
#include <xlnt/cell/cell.hpp>
#include <xlnt/styles/alignment.hpp>
#include <xlnt/styles/border.hpp>
#include <xlnt/styles/fill.hpp>
#include <xlnt/styles/font.hpp>
#include <xlnt/styles/format.hpp>
#include <xlnt/styles/style.hpp>
#include <xlnt/styles/number_format.hpp>
#include <xlnt/styles/protection.hpp>
#include <xlnt/workbook/worksheet_iterator.hpp>
#include <xlnt/worksheet/cell_iterator.hpp>
#include <xlnt/worksheet/cell_vector.hpp>
#include <xlnt/worksheet/range.hpp>
#include <xlnt/worksheet/worksheet.hpp>
#include <xlnt/worksheet/range_iterator.hpp>

namespace {

// Miscellaneous Functions

struct EnumClassHash
{
    template <typename T>
    std::size_t operator()(T t) const
    {
        return static_cast<std::size_t>(t);
    }
};

std::string string_lower(std::string str)
{
    for (std::size_t i = 0; i < str.size(); i++)
    {
        str[i] = std::tolower(str[i]);
    }
    
    return str;
}

bool is_true(const std::string &bool_string)
{
    return bool_string == "1" || bool_string == "true";
}

bool is_false(const std::string &bool_string)
{
    return bool_string == "0" || bool_string == "false";
}

std::size_t string_to_size_t(const std::string &s)
{
#if ULLONG_MAX == SIZE_MAX
	return std::stoull(s);
#else
	return std::stoul(s);
#endif
}

//
// enum serialization
//

// font::underline_style serialization

const std::unordered_map<std::string, xlnt::font::underline_style> &get_string_underline_style_map()
{
    static std::unordered_map<std::string, xlnt::font::underline_style> *map = nullptr;
    
    if (map == nullptr)
    {
        map = new std::unordered_map<std::string, xlnt::font::underline_style>
        {
            { "double", xlnt::font::underline_style::double_ },
            { "double-accounting", xlnt::font::underline_style::double_accounting },
            { "none", xlnt::font::underline_style::none },
            { "single", xlnt::font::underline_style::single },
            { "single-accounting", xlnt::font::underline_style::single_accounting }
        };
    }
    
    return *map;
}

const std::unordered_map<xlnt::font::underline_style, std::string, EnumClassHash> &get_underline_style_string_map()
{
    static std::unordered_map<xlnt::font::underline_style, std::string, EnumClassHash> *map = nullptr;
    
    if (map == nullptr)
    {
        map = new std::unordered_map<xlnt::font::underline_style, std::string, EnumClassHash>;
        
        for (auto pair : get_string_underline_style_map())
        {
            map->emplace(pair.second, pair.first);
        }
    }
    
    return *map;
}

xlnt::font::underline_style underline_style_from_string(const std::string &underline_string)
{
    return get_string_underline_style_map().at(string_lower(underline_string));
}

std::string underline_style_to_string(xlnt::font::underline_style underline_style)
{
    return get_underline_style_string_map().at(underline_style);
}


// fill::pattern_type serialization

const std::unordered_map<std::string, xlnt::fill::pattern_type> &get_string_pattern_fill_type_map()
{
    static std::unordered_map<std::string, xlnt::fill::pattern_type> *map = nullptr;
    
    if (map == nullptr)
    {
        map = new std::unordered_map<std::string, xlnt::fill::pattern_type>
        {
            { "darkdown", xlnt::fill::pattern_type::darkdown },
            { "darkgray", xlnt::fill::pattern_type::darkgray },
            { "darkgrid", xlnt::fill::pattern_type::darkgrid },
            { "darkhorizontal", xlnt::fill::pattern_type::darkhorizontal },
            { "darktrellis", xlnt::fill::pattern_type::darktrellis },
            { "darkup", xlnt::fill::pattern_type::darkup },
            { "darkvertical", xlnt::fill::pattern_type::darkvertical },
            { "gray0625", xlnt::fill::pattern_type::gray0625 },
            { "gray125", xlnt::fill::pattern_type::gray125 },
            { "lightdown", xlnt::fill::pattern_type::lightdown },
            { "lightgray", xlnt::fill::pattern_type::lightgray },
            { "lightgrid", xlnt::fill::pattern_type::lightgrid },
            { "lighthorizontal", xlnt::fill::pattern_type::lighthorizontal },
            { "lighttrellis", xlnt::fill::pattern_type::lighttrellis },
            { "lightup", xlnt::fill::pattern_type::lightup },
            { "lightvertical", xlnt::fill::pattern_type::lightvertical },
            { "mediumgray", xlnt::fill::pattern_type::mediumgray },
            { "none", xlnt::fill::pattern_type::none },
            { "solid", xlnt::fill::pattern_type::solid }
        };
    }
    
    return *map;
}

const std::unordered_map<xlnt::fill::pattern_type, std::string, EnumClassHash> &get_pattern_fill_type_string_map()
{
    static std::unordered_map<xlnt::fill::pattern_type, std::string, EnumClassHash> *map = nullptr;
    
    if (map == nullptr)
    {
        map = new std::unordered_map<xlnt::fill::pattern_type, std::string, EnumClassHash>;

        for (auto pair : get_string_pattern_fill_type_map())
        {
            map->emplace(pair.second, pair.first);
        }
    }
    
    return *map;
}

xlnt::fill::pattern_type pattern_fill_type_from_string(const std::string &fill_type)
{
    return get_string_pattern_fill_type_map().at(string_lower(fill_type));
};

std::string pattern_fill_type_to_string(xlnt::fill::pattern_type fill_type)
{
    return get_pattern_fill_type_string_map().at(fill_type);
}


// border_style serialization

const std::unordered_map<std::string, xlnt::border_style> &get_string_border_style_map()
{
    static std::unordered_map<std::string, xlnt::border_style> *map = nullptr;
    
    if (map == nullptr)
    {
        map = new std::unordered_map<std::string, xlnt::border_style>
        {
            { "dashdot", xlnt::border_style::dashdot },
            { "dashdotdot", xlnt::border_style::dashdotdot },
            { "dashed", xlnt::border_style::dashed },
            { "dotted", xlnt::border_style::dotted },
            { "double", xlnt::border_style::double_ },
            { "hair", xlnt::border_style::hair },
            { "medium", xlnt::border_style::medium },
            { "mediumdashdot", xlnt::border_style::mediumdashdot },
            { "mediumdashdotdot", xlnt::border_style::mediumdashdotdot },
            { "mediumdashed", xlnt::border_style::mediumdashed },
            { "none", xlnt::border_style::none },
            { "slantdashdot", xlnt::border_style::slantdashdot },
            { "thick", xlnt::border_style::thick },
            { "thin", xlnt::border_style::thin }
        };
    }
    
    return *map;
}

const std::unordered_map<xlnt::border_style, std::string, EnumClassHash> &get_border_style_string_map()
{
    static std::unordered_map<xlnt::border_style, std::string, EnumClassHash> *map = nullptr;
    
    if (map == nullptr)
    {
        map = new std::unordered_map<xlnt::border_style, std::string, EnumClassHash>;

        for (auto pair : get_string_border_style_map())
        {
            map->emplace(pair.second, pair.first);
        }
    }
    
    return *map;
}

xlnt::border_style border_style_from_string(const std::string &border_style_string)
{
    return get_string_border_style_map().at(string_lower(border_style_string));
}

std::string border_style_to_string(xlnt::border_style border_style)
{
    return get_border_style_string_map().at(border_style);
}


// vertical_alignment serialization

const std::unordered_map<std::string, xlnt::vertical_alignment> &get_string_vertical_alignment_map()
{
    static std::unordered_map<std::string, xlnt::vertical_alignment> *map = nullptr;
    
    if (map == nullptr)
    {
        map = new std::unordered_map<std::string, xlnt::vertical_alignment>
        {
            { "bottom", xlnt::vertical_alignment::bottom },
            { "center", xlnt::vertical_alignment::center },
            { "justify", xlnt::vertical_alignment::justify },
            { "none", xlnt::vertical_alignment::none },
            { "top", xlnt::vertical_alignment::top }
        };
    }
    
    return *map;
}

const std::unordered_map<xlnt::vertical_alignment, std::string, EnumClassHash> &get_vertical_alignment_string_map()
{
    static std::unordered_map<xlnt::vertical_alignment, std::string, EnumClassHash> *map = nullptr;
    
    if (map == nullptr)
    {
        map = new std::unordered_map<xlnt::vertical_alignment, std::string, EnumClassHash>;

        for (auto pair : get_string_vertical_alignment_map())
        {
            map->emplace(pair.second, pair.first);
        }
    }
    
    return *map;
}

xlnt::vertical_alignment vertical_alignment_from_string(const std::string &vertical_alignment_string)
{
    return get_string_vertical_alignment_map().at(string_lower(vertical_alignment_string));
}

std::string vertical_alignment_to_string(xlnt::vertical_alignment vertical_alignment)
{
    return get_vertical_alignment_string_map().at(vertical_alignment);
}

// horizontal_alignment


const std::unordered_map<std::string, xlnt::horizontal_alignment> &get_string_horizontal_alignment_map()
{
    static std::unordered_map<std::string, xlnt::horizontal_alignment> *map = nullptr;
    
    if (map == nullptr)
    {
        map = new std::unordered_map<std::string, xlnt::horizontal_alignment>
        {
            { "center", xlnt::horizontal_alignment::center },
            { "center-continous", xlnt::horizontal_alignment::center_continuous },
            { "general", xlnt::horizontal_alignment::general },
            { "justify", xlnt::horizontal_alignment::justify },
            { "left", xlnt::horizontal_alignment::left },
            { "none", xlnt::horizontal_alignment::none },
            { "right", xlnt::horizontal_alignment::right }
        };
    }
    
    return *map;
}

const std::unordered_map<xlnt::horizontal_alignment, std::string, EnumClassHash> &get_horizontal_alignment_string_map()
{
    static std::unordered_map<xlnt::horizontal_alignment, std::string, EnumClassHash> *map = nullptr;
    
    if (map == nullptr)
    {
        map = new std::unordered_map<xlnt::horizontal_alignment, std::string, EnumClassHash>;

        for (auto pair : get_string_horizontal_alignment_map())
        {
            map->emplace(pair.second, pair.first);
        }
    }
    
    return *map;
}

xlnt::horizontal_alignment horizontal_alignment_from_string(const std::string &horizontal_alignment_string)
{
    return get_string_horizontal_alignment_map().at(string_lower(horizontal_alignment_string));
}

std::string horizontal_alignment_to_string(xlnt::horizontal_alignment horizontal_alignment)
{
    return get_horizontal_alignment_string_map().at(horizontal_alignment);
}

// Reading

xlnt::protection read_protection(const pugi::xml_node protection_node)
{
    xlnt::protection prot;

    if (is_true(protection_node.attribute("locked").value()))
    {
        prot.set_locked(true);
    }
    else if (!is_false(protection_node.attribute("locked").value()))
    {
        throw std::runtime_error("bad protection value");
    }

    if (is_true(protection_node.attribute("hidden").value()))
    {
        prot.set_hidden(true);
    }
    else if (!is_false(protection_node.attribute("hidden").value()))
    {
        throw std::runtime_error("bad protection value");
    }

    return prot;
}

xlnt::alignment read_alignment(const pugi::xml_node alignment_node)
{
    xlnt::alignment align;

    align.set_wrap_text(is_true(alignment_node.attribute("wrapText").value()));
    align.set_shrink_to_fit(is_true(alignment_node.attribute("shrinkToFit").value()));

    if (alignment_node.attribute("vertical"))
    {
        std::string vertical = alignment_node.attribute("vertical").value();
        align.set_vertical(vertical_alignment_from_string(vertical));
    }

    if (alignment_node.attribute("horizontal"))
    {
        std::string horizontal = alignment_node.attribute("horizontal").value();
        align.set_horizontal(horizontal_alignment_from_string(horizontal));
    }

    return align;
}

void read_number_formats(const pugi::xml_node number_formats_node, std::vector<xlnt::number_format> &number_formats)
{
    number_formats.clear();

    for (auto num_fmt_node : number_formats_node.children("numFmt"))
    {
        std::string format_string(num_fmt_node.attribute("formatCode").value());

        if (format_string == "GENERAL")
        {
            format_string = "General";
        }

        xlnt::number_format nf;

        nf.set_format_string(format_string);
        nf.set_id(string_to_size_t(num_fmt_node.attribute("numFmtId").value()));

        number_formats.push_back(nf);
    }
}

xlnt::color read_color(const pugi::xml_node &color_node)
{
    if (color_node.attribute("rgb"))
    {
        return xlnt::color(xlnt::color::type::rgb, color_node.attribute("rgb").value());
    }
    else if (color_node.attribute("theme"))
    {
        return xlnt::color(xlnt::color::type::theme, string_to_size_t(color_node.attribute("theme").value()));
    }
    else if (color_node.attribute("indexed"))
    {
        return xlnt::color(xlnt::color::type::indexed, string_to_size_t(color_node.attribute("indexed").value()));
    }
    else if (color_node.attribute("auto"))
    {
        return xlnt::color(xlnt::color::type::auto_, string_to_size_t(color_node.attribute("auto").value()));
    }

    throw std::runtime_error("bad color");
}

xlnt::font read_font(const pugi::xml_node font_node)
{
    xlnt::font new_font;

    new_font.set_size(string_to_size_t(font_node.child("sz").attribute("val").value()));
    new_font.set_name(font_node.child("name").attribute("val").value());

    if (font_node.child("color"))
    {
        new_font.set_color(read_color(font_node.child("color")));
    }

    if (font_node.child("family"))
    {
        new_font.set_family(string_to_size_t(font_node.child("family").attribute("val").value()));
    }

    if (font_node.child("scheme"))
    {
        new_font.set_scheme(font_node.child("scheme").attribute("val").value());
    }

    if (font_node.child("b"))
    {
        if(font_node.child("b").attribute("val"))
        {
            new_font.set_bold(is_true(font_node.child("b").attribute("val").value()));
        }
        else
        {
            new_font.set_bold(true);
        }
    }

    if (font_node.child("strike"))
    {
        if(font_node.child("strike").attribute("val"))
        {
            new_font.set_strikethrough(is_true(font_node.child("strike").attribute("val").value()));
        }
        else
        {
            new_font.set_strikethrough(true);
        }
    }

    if (font_node.child("i"))
    {
        if(font_node.child("i").attribute("val"))
        {
            new_font.set_italic(is_true(font_node.child("i").attribute("val").value()));
        }
        else
        {
            new_font.set_italic(true);
        }
    }

    if (font_node.child("u"))
    {
        if (font_node.child("u").attribute("val"))
        {
            std::string underline_string = font_node.child("u").attribute("val").value();
            new_font.set_underline(underline_style_from_string(underline_string));
        }
        else
        {
            new_font.set_underline(xlnt::font::underline_style::single);
        }
    }

    return new_font;
}

void read_fonts(const pugi::xml_node &fonts_node, std::vector<xlnt::font> &fonts)
{
    fonts.clear();

    for (auto font_node : fonts_node.children())
    {
        fonts.push_back(read_font(font_node));
    }
}

void read_indexed_colors(const pugi::xml_node &indexed_colors_node, std::vector<xlnt::color> &colors)
{
    for (auto color_node : indexed_colors_node.children())
    {
        colors.push_back(read_color(color_node));
    }
}

void read_colors(const pugi::xml_node &colors_node, std::vector<xlnt::color> &colors)
{
    colors.clear();

    if (colors_node.child("indexedColors"))
    {
        read_indexed_colors(colors_node.child("indexedColors"), colors);
    }
}

xlnt::fill read_fill(const pugi::xml_node &fill_node)
{
    xlnt::fill new_fill;

    if (fill_node.child("patternFill"))
    {
        new_fill.set_type(xlnt::fill::type::pattern);

        auto pattern_fill_node = fill_node.child("patternFill");
        std::string pattern_fill_type_string = pattern_fill_node.attribute("patternType").value();
        
        if (!pattern_fill_type_string.empty())
        {
            new_fill.set_pattern_type(pattern_fill_type_from_string(pattern_fill_type_string));

            if (pattern_fill_node.child("bgColor"))
            {
                new_fill.get_background_color() = read_color(pattern_fill_node.child("bgColor"));
            }

            if (pattern_fill_node.child("fgColor"))
            {
                new_fill.get_foreground_color() = read_color(pattern_fill_node.child("fgColor"));
            }
        }
    }

    return new_fill;
}

void read_fills(const pugi::xml_node &fills_node, std::vector<xlnt::fill> &fills)
{
    fills.clear();

    for (auto fill_node : fills_node.children())
    {
        fills.emplace_back();
        fills.back() = read_fill(fill_node);
    }
}

xlnt::side read_side(const pugi::xml_node &side_node)
{
    xlnt::side new_side;

    if (side_node.attribute("style"))
    {
        new_side.get_border_style() = border_style_from_string(side_node.attribute("style").value());
    }

    if (side_node.child("color"))
    {
        new_side.get_color() = read_color(side_node.child("color"));
    }

    return new_side;
}

xlnt::border read_border(const pugi::xml_node &border_node)
{
    xlnt::border new_border;

    if (border_node.child("start"))
    {
        new_border.get_start() = read_side(border_node.child("start"));
    }

    if (border_node.child("end"))
    {
        new_border.get_end() = read_side(border_node.child("end"));
    }

    if (border_node.child("left"))
    {
        new_border.get_left() = read_side(border_node.child("left"));
    }

    if (border_node.child("right"))
    {
        new_border.get_right() = read_side(border_node.child("right"));
    }

    if (border_node.child("top"))
    {
        new_border.get_top() = read_side(border_node.child("top"));
    }

    if (border_node.child("bottom"))
    {
        new_border.get_bottom() = read_side(border_node.child("bottom"));
    }

    if (border_node.child("diagonal"))
    {
        new_border.get_diagonal() = read_side(border_node.child("diagonal"));
    }

    if (border_node.child("vertical"))
    {
        new_border.get_vertical() = read_side(border_node.child("vertical"));
    }

    if (border_node.child("horizontal"))
    {
        new_border.get_horizontal() = read_side(border_node.child("horizontal"));
    }

    return new_border;
}

void read_borders(const pugi::xml_node &borders_node, std::vector<xlnt::border> &borders)
{
    borders.clear();

    for (auto border_node : borders_node.children())
    {
        borders.push_back(read_border(border_node));
    }
}


bool read_base_format(const pugi::xml_node &format_node, const xlnt::detail::stylesheet &stylesheet, xlnt::base_format &f)
{
    // Alignment
    f.alignment_applied(format_node.child("alignment") || is_true(format_node.attribute("applyAlignment").value()));

    if (f.alignment_applied())
    {
        auto inline_alignment = read_alignment(format_node.child("alignment"));
        f.set_alignment(inline_alignment);
    }
    
    // Border
    auto border_index = format_node.attribute("borderId") ? string_to_size_t(format_node.attribute("borderId").value()) : 0;
    f.set_border(stylesheet.borders.at(border_index));
    f.border_applied(is_true(format_node.attribute("applyBorder").value()));
    
    // Fill
    auto fill_index = format_node.attribute("fillId") ? string_to_size_t(format_node.attribute("fillId").value()) : 0;
    f.set_fill(stylesheet.fills.at(fill_index));
    f.fill_applied(is_true(format_node.attribute("applyFill").value()));
    
    // Font
    auto font_index = format_node.attribute("fontId") ? string_to_size_t(format_node.attribute("fontId").value()) : 0;
    f.set_font(stylesheet.fonts.at(font_index));
    f.font_applied(is_true(format_node.attribute("applyFont").value()));

    // Number Format
    auto number_format_id = string_to_size_t(format_node.attribute("numFmtId").value());

    bool builtin_format = true;

    for (const auto &num_fmt : stylesheet.number_formats)
    {
        if (static_cast<std::size_t>(num_fmt.get_id()) == number_format_id)
        {
            f.set_number_format(num_fmt);
            builtin_format = false;
            break;
        }
    }

    if (builtin_format)
    {
        f.set_number_format(xlnt::number_format::from_builtin_id(number_format_id));
    }
    
    f.number_format_applied(is_true(format_node.attribute("applyNumberFormat").value()));

    // Protection
    f.protection_applied(format_node.attribute("protection") || is_true(format_node.attribute("applyProtection").value()));

    if (f.protection_applied())
    {
        auto inline_protection = read_protection(format_node.child("protection"));
        f.set_protection(inline_protection);
    }
    
    return true;
}


void read_formats(const pugi::xml_node &formats_node, const xlnt::detail::stylesheet &stylesheet,
    std::vector<xlnt::format> &formats, std::vector<std::string> &format_styles)
{
    for (auto format_node : formats_node.children("xf"))
    {
        xlnt::format format;
        read_base_format(format_node, stylesheet, format);
        
        // TODO do all formats have xfId?
        if(format_node.attribute("xfId"))
        {
            auto style_index = string_to_size_t(format_node.attribute("xfId").value());
            auto style_name = stylesheet.style_name_map.at(style_index);
            format_styles.push_back(style_name);
        }
        else
        {
            format_styles.push_back("");
        }
        
        formats.push_back(format);
    }
}

xlnt::style read_style(const pugi::xml_node &style_node, const pugi::xml_node &style_format_node, const xlnt::detail::stylesheet &stylesheet)
{
    xlnt::style s;

    read_base_format(style_format_node, stylesheet, s);

    s.set_name(style_node.attribute("name").value());
    s.set_hidden(style_node.attribute("hidden") && is_true(style_node.attribute("hidden").value()));
    s.set_builtin_id(string_to_size_t(style_node.attribute("builtinId").value()));
    
    return s;
}

void read_styles(const pugi::xml_node &styles_node, const pugi::xml_node &style_formats_node, const xlnt::detail::stylesheet stylesheet, std::vector<xlnt::style> &styles, std::unordered_map<std::size_t, std::string> &style_names)
{
    std::size_t style_index = 0;
    
    for (auto cell_style_format_node : style_formats_node.children())
    {
        bool match = false;
        
        for (auto cell_style_node : styles_node.children())
        {
            auto cell_style_format_index = std::stoull(cell_style_node.attribute("xfId").value());
            
            if (cell_style_format_index == style_index)
            {
                styles.push_back(read_style(cell_style_node, cell_style_format_node, stylesheet));
                style_names[style_index] = styles.back().get_name();
                match = true;
                
                break;
            }
        }

        style_index++;
    }
}

bool write_color(const xlnt::color &color, pugi::xml_node color_node)
{
    switch (color.get_type())
    {
    case xlnt::color::type::auto_:
        color_node.append_attribute("auto").set_value(std::to_string(color.get_auto()).c_str());
        break;
    case xlnt::color::type::theme:
        color_node.append_attribute("theme").set_value(std::to_string(color.get_theme()).c_str());
        break;
    case xlnt::color::type::indexed:
        color_node.append_attribute("indexed").set_value(std::to_string(color.get_index()).c_str());
        break;
    case xlnt::color::type::rgb:
        color_node.append_attribute("rgb").set_value(color.get_rgb_string().c_str());
        break;
    default:
        throw std::runtime_error("bad type");
    }
    
    return true;
}

bool write_fonts(const std::vector<xlnt::font> &fonts, pugi::xml_node &fonts_node)
{
    fonts_node.append_attribute("count").set_value(std::to_string(fonts.size()).c_str());
    // TODO: what does this do?
    // fonts_node.append_attribute("x14ac:knownFonts", "1");

    for (auto &f : fonts)
    {
        auto font_node = fonts_node.append_child("font");

        if (f.is_bold())
        {
            auto bold_node = font_node.append_child("b");
            bold_node.append_attribute("val").set_value("1");
        }

        if (f.is_italic())
        {
            auto italic_node = font_node.append_child("i");
            italic_node.append_attribute("val").set_value("1");
        }

        if (f.is_underline())
        {
            auto underline_node = font_node.append_child("u");
            underline_node.append_attribute("val").set_value(underline_style_to_string(f.get_underline()).c_str());
        }

        if (f.is_strikethrough())
        {
            auto strike_node = font_node.append_child("strike");
            strike_node.append_attribute("val").set_value("1");
        }

        auto size_node = font_node.append_child("sz");
        size_node.append_attribute("val").set_value(std::to_string(f.get_size()).c_str());

        auto color_node = font_node.append_child("color");
        
        write_color(f.get_color(), color_node);

        auto name_node = font_node.append_child("name");
        name_node.append_attribute("val").set_value(f.get_name().c_str());

        if (f.has_family())
        {
            auto family_node = font_node.append_child("family");
            family_node.append_attribute("val").set_value(std::to_string(f.get_family()).c_str());
        }

        if (f.has_scheme())
        {
            auto scheme_node = font_node.append_child("scheme");
            scheme_node.append_attribute("val").set_value("minor");
        }
    }
    
    return true;
}

bool write_fills(const std::vector<xlnt::fill> &fills, pugi::xml_node &fills_node)
{
    fills_node.append_attribute("count").set_value(std::to_string(fills.size()).c_str());

    for (auto &fill_ : fills)
    {
        auto fill_node = fills_node.append_child("fill");

        if (fill_.get_type() == xlnt::fill::type::pattern)
        {
            auto pattern_fill_node = fill_node.append_child("patternFill");
            pattern_fill_node.append_attribute("patternType").set_value(pattern_fill_type_to_string(fill_.get_pattern_type()).c_str());
            
            if (fill_.get_pattern_type() != xlnt::fill::pattern_type::solid) continue;

            if (fill_.get_foreground_color())
            {
                write_color(*fill_.get_foreground_color(), pattern_fill_node.append_child("fgColor"));
            }

            if (fill_.get_background_color())
            {
                write_color(*fill_.get_background_color(), pattern_fill_node.append_child("bgColor"));
            }
        }
        else if (fill_.get_type() == xlnt::fill::type::solid)
        {
            auto solid_fill_node = fill_node.append_child("solidFill");
            solid_fill_node.append_child("color");
        }
        else if (fill_.get_type() == xlnt::fill::type::gradient)
        {
            auto gradient_fill_node = fill_node.append_child("gradientFill");

            if (fill_.get_gradient_type() == xlnt::fill::gradient_type::linear)
            {
                gradient_fill_node.append_attribute("degree").set_value(std::to_string(fill_.get_rotation()).c_str());
            }
            else if (fill_.get_gradient_type() == xlnt::fill::gradient_type::path)
            {
                gradient_fill_node.append_attribute("left").set_value(std::to_string(fill_.get_gradient_left()).c_str());
                gradient_fill_node.append_attribute("right").set_value(std::to_string(fill_.get_gradient_right()).c_str());
                gradient_fill_node.append_attribute("top").set_value(std::to_string(fill_.get_gradient_top()).c_str());
                gradient_fill_node.append_attribute("bottom").set_value(std::to_string(fill_.get_gradient_bottom()).c_str());

                auto start_node = gradient_fill_node.append_child("stop");
                start_node.append_attribute("position").set_value("0");

                auto end_node = gradient_fill_node.append_child("stop");
                end_node.append_attribute("position").set_value("1");
            }
        }
    }

    return true;
}

bool write_borders(const std::vector<xlnt::border> &borders, pugi::xml_node &borders_node)
{
    borders_node.append_attribute("count").set_value(std::to_string(borders.size()).c_str());

    for (const auto &border_ : borders)
    {
        auto border_node = borders_node.append_child("border");

        std::vector<std::tuple<std::string, const std::experimental::optional<xlnt::side>>> sides;
        
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
                auto side_node = border_node.append_child(current_name.c_str());

                if (current_side->get_border_style())
                {
                    auto style_string = border_style_to_string(*current_side->get_border_style());
                    side_node.append_attribute("style").set_value(style_string.c_str());
                }

                if (current_side->get_color())
                {
                    auto color_node = side_node.append_child("color");
                    write_color(*current_side->get_color(), color_node);
                }
            }
        }
    }

    return true;
}

bool write_base_format(const xlnt::base_format &xf, const xlnt::detail::stylesheet &stylesheet, pugi::xml_node xf_node)
{
    xf_node.append_attribute("numFmtId").set_value(std::to_string(xf.get_number_format().get_id()).c_str());
     
    auto font_id = std::distance(stylesheet.fonts.begin(), std::find(stylesheet.fonts.begin(), stylesheet.fonts.end(), xf.get_font()));
    xf_node.append_attribute("fontId").set_value(std::to_string(font_id).c_str());

    auto fill_id = std::distance(stylesheet.fills.begin(), std::find(stylesheet.fills.begin(), stylesheet.fills.end(), xf.get_fill()));
    xf_node.append_attribute("fillId").set_value(std::to_string(fill_id).c_str());

    auto border_id = std::distance(stylesheet.borders.begin(), std::find(stylesheet.borders.begin(), stylesheet.borders.end(), xf.get_border()));
    xf_node.append_attribute("borderId").set_value(std::to_string(border_id).c_str());

    if(xf.number_format_applied()) xf_node.append_attribute("applyNumberFormat").set_value("1");
    if(xf.fill_applied()) xf_node.append_attribute("applyFill").set_value("1");
    if(xf.font_applied()) xf_node.append_attribute("applyFont").set_value("1");
    if(xf.border_applied()) xf_node.append_attribute("applyBorder").set_value("1");
    if(xf.alignment_applied()) xf_node.append_attribute("applyAlignment").set_value("1");
    if(xf.protection_applied()) xf_node.append_attribute("applyProtection").set_value("1");

    if (xf.alignment_applied())
    {
        auto alignment_node = xf_node.append_child("alignment");

        if (xf.get_alignment().has_vertical())
        {
            auto vertical = vertical_alignment_to_string(xf.get_alignment().get_vertical());
            alignment_node.append_attribute("vertical").set_value(vertical.c_str());
        }

        if (xf.get_alignment().has_horizontal())
        {
            auto horizontal = horizontal_alignment_to_string(xf.get_alignment().get_horizontal());
            alignment_node.append_attribute("horizontal").set_value(horizontal.c_str());
        }

        if (xf.get_alignment().get_wrap_text())
        {
            alignment_node.append_attribute("wrapText").set_value("1");
        }
        
        if (xf.get_alignment().get_shrink_to_fit())
        {
            alignment_node.append_attribute("shrinkToFit").set_value("1");
        }
    }
    
    if (xf.protection_applied())
    {
        auto protection_node = xf_node.append_child("protection");
        protection_node.append_attribute("locked").set_value(xf.get_protection().get_locked() ? "1" : "0");
        protection_node.append_attribute("hidden").set_value(xf.get_protection().get_hidden() ? "1" : "0");
    }
    
    return true;
}

bool write_styles(const xlnt::detail::stylesheet &stylesheet, pugi::xml_node &styles_node, pugi::xml_node &style_formats_node)
{
    style_formats_node.append_attribute("count").set_value(std::to_string(stylesheet.styles.size()).c_str());
    styles_node.append_attribute("count").set_value(std::to_string(stylesheet.styles.size()).c_str());
    std::size_t style_index = 0;

    for(auto &current_style : stylesheet.styles)
    {
        auto xf_node = style_formats_node.append_child("xf");
        write_base_format(current_style, stylesheet, xf_node);

        auto cell_style_node = styles_node.append_child("cellStyle");
        
        cell_style_node.append_attribute("name").set_value(current_style.get_name().c_str());
        cell_style_node.append_attribute("xfId").set_value(std::to_string(style_index++).c_str());
        cell_style_node.append_attribute("builtinId").set_value(std::to_string(current_style.get_builtin_id()).c_str());
        
        if (current_style.get_hidden())
        {
            cell_style_node.append_attribute("hidden").set_value("1");
        }
    }

    return true;
}

bool write_formats(const xlnt::detail::stylesheet &stylesheet, pugi::xml_node &formats_node)
{
    formats_node.append_attribute("count").set_value(std::to_string(stylesheet.formats.size()).c_str());

    auto format_style_iterator = stylesheet.format_styles.begin();

    for(auto &current_format : stylesheet.formats)
    {
        auto xf_node = formats_node.append_child("xf");
        write_base_format(current_format, stylesheet, xf_node);
        
        const auto format_style_name = *(format_style_iterator++);
        
        if(!format_style_name.empty())
        {
            auto style = std::find_if(stylesheet.styles.begin(), stylesheet.styles.end(),
                [&](const xlnt::style &s) { return s.get_name() == format_style_name; });
            auto style_index = std::distance(stylesheet.styles.begin(), style);
            
            xf_node.append_attribute("xfId").set_value(std::to_string(style_index).c_str());
        }
    }

    return true;
}

bool write_dxfs(pugi::xml_node &dxfs_node)
{
    dxfs_node.append_attribute("count").set_value("0");
    return true;
}

bool write_table_styles(pugi::xml_node &table_styles_node)
{
    table_styles_node.append_attribute("count").set_value("0");
    table_styles_node.append_attribute("defaultTableStyle").set_value("TableStyleMedium9");
    table_styles_node.append_attribute("defaultPivotStyle").set_value("PivotStyleMedium7");
    
    return true;
}

bool write_colors(const std::vector<xlnt::color> &colors, pugi::xml_node &colors_node)
{
    auto indexed_colors_node = colors_node.append_child("indexedColors");

    for (auto &c : colors)
    {
        indexed_colors_node.append_child("rgbColor").append_attribute("rgb").set_value(c.get_rgb_string().c_str());
    }
    
    return true;
}

bool write_number_formats(const std::vector<xlnt::number_format> &number_formats, pugi::xml_node &number_formats_node)
{
    number_formats_node.append_attribute("count").set_value(std::to_string(number_formats.size()).c_str());

    for (const auto &num_fmt : number_formats)
    {
        auto num_fmt_node = number_formats_node.append_child("numFmt");
        num_fmt_node.append_attribute("numFmtId").set_value(std::to_string(num_fmt.get_id()).c_str());
        num_fmt_node.append_attribute("formatCode").set_value(num_fmt.get_format_string().c_str());
    }

    return true;
}

} // namespace

namespace xlnt {

style_serializer::style_serializer(detail::stylesheet &stylesheet) : stylesheet_(stylesheet)
{
}

bool style_serializer::read_stylesheet(const pugi::xml_document &xml)
{
    auto stylesheet_node = xml.child("styleSheet");

    read_borders(stylesheet_node.child("borders"), stylesheet_.borders);
    read_fills(stylesheet_node.child("fills"), stylesheet_.fills);
    read_fonts(stylesheet_node.child("fonts"), stylesheet_.fonts);
    read_number_formats(stylesheet_node.child("numFmts"), stylesheet_.number_formats);
    read_colors(stylesheet_node.child("colors"), stylesheet_.colors);
    read_styles(stylesheet_node.child("cellStyles"), stylesheet_node.child("cellStyleXfs"), stylesheet_, stylesheet_.styles, stylesheet_.style_name_map);
    read_formats(stylesheet_node.child("cellXfs"), stylesheet_, stylesheet_.formats, stylesheet_.format_styles);

    return true;
}

bool style_serializer::write_stylesheet(pugi::xml_document &doc)
{
    auto root_node = doc.append_child("styleSheet");
    root_node.append_attribute("xmlns").set_value("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    root_node.append_attribute("xmlns:mc").set_value("http://schemas.openxmlformats.org/markup-compatibility/2006");
    root_node.append_attribute("mc:Ignorable").set_value("x14ac");
    root_node.append_attribute("xmlns:x14ac").set_value("http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

    if (!stylesheet_.number_formats.empty())
    {
        auto number_formats_node = root_node.append_child("numFmts");
        write_number_formats(stylesheet_.number_formats, number_formats_node);
    }

    if (!stylesheet_.fonts.empty())
    {
        auto fonts_node = root_node.append_child("fonts");
        write_fonts(stylesheet_.fonts, fonts_node);
    }

    if (!stylesheet_.fills.empty())
    {
        auto fills_node = root_node.append_child("fills");
        write_fills(stylesheet_.fills, fills_node);
    }

    if (!stylesheet_.borders.empty())
    {
        auto borders_node = root_node.append_child("borders");
        write_borders(stylesheet_.borders, borders_node);
    }
    
    auto cell_style_xfs_node = root_node.append_child("cellStyleXfs");
    
    auto cell_xfs_node = root_node.append_child("cellXfs");
    write_formats(stylesheet_, cell_xfs_node);

    auto cell_styles_node = root_node.append_child("cellStyles");
    write_styles(stylesheet_, cell_styles_node, cell_style_xfs_node);
    
    auto dxfs_node = root_node.append_child("dxfs");
    write_dxfs(dxfs_node);

    auto table_styles_node = root_node.append_child("tableStyles");
    write_table_styles(table_styles_node);

    if(!stylesheet_.colors.empty())
    {
        auto colors_node = root_node.append_child("colors");
        write_colors(stylesheet_.colors, colors_node);
    }

    return true;
}

} // namespace xlnt
