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
#pragma once

#include <string>
#include <unordered_map>
#include <vector>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/workbook/workbook.hpp>

namespace xlnt {

class alignment;
class border;
class color;
class conditional_format;
class fill;
class font;
class named_style;
class number_format;
class protection;
class side;
class style;
class xml_document;
class xml_node;

/// <summary>
/// Reads and writes xl/styles.xml from/to an associated workbook.
/// </summary>
class XLNT_CLASS style_serializer
{
  public:
    /// <summary>
    /// Construct a style_serializer which can write styles.xml based on wb or populate wb
    /// with styles from an existing styles.xml.
    style_serializer(workbook &wb);

    //
    // Primary methods
    //

    /// <summary>
    /// Load all styles from xml_document into workbook given in constructor.
    /// </summary>
    bool read_stylesheet(const xml_document &xml);

    /// <summary>
    /// Populate parameter xml with an XML tree representing the styles contained in the workbook
    /// given in the constructor.
    /// </summary>
    xml_document write_stylesheet() const;

    // TODO: These need to be public for unit tests. Could also make test class a friend?
    // private:

    //
    // Static element readers (i.e. readers that don't use internal state)
    //

    /// <summary>
    /// Read and return an alignment from alignment_node.
    /// </summary>
    static alignment read_alignment(const xml_node &alignment_node);

    /// <summary>
    /// Read and return a border side from side_node.
    /// </summary>
    static side read_side(const xml_node &side_node);

    /// <summary>
    /// Read and return a border from border_node.
    /// </summary>
    static border read_border(const xml_node &border_node);

    /// <summary>
    /// Read and return a fill from fill_node.
    /// </summary>
    static fill read_fill(const xml_node &fill_node);

    /// <summary>
    /// Read and return a font from font_node.
    /// </summary>
    static font read_font(const xml_node &font_node);

    /// <summary>
    /// Read and return a number format from number_format_node.
    /// </summary>
    static number_format read_number_format(const xml_node &number_format_node);

    /// <summary>
    /// Read and return a protection from protection_node.
    /// </summary>
    static protection read_protection(const xml_node &protection_node);

    /// <summary>
    /// Read and return all indexed colors from indexed_colors_node.
    /// </summary>
    static std::vector<color> read_indexed_colors(const xml_node &indexed_colors_node);

    /// <summary>
    /// Read and return a color from color_node.
    /// </summary>
    static color read_color(const xml_node &color_node);

    /// <summary>
    /// Read and return a conditional format, dxf, from conditional_format_node.
    /// </summary>
    static conditional_format read_conditional_format(const xml_node &conditional_formats_node);

    /// <summary>
    /// Read and return a pair containing a name and corresponding style from named_style_node.
    /// </summary>
    static std::pair<std::string, style> read_named_style(const xml_node &named_style_node);

    //
    // Static element writers (i.e. writers that don't use internal state)
    //

    /// <summary>
    /// Build and return an xml tree representing alignment_.
    /// </summary>
    static xml_node write_alignment(const alignment &alignment_);

    /// <summary>
    /// Build and return an xml tree representing border_.
    /// </summary>
    static xml_node write_border(const border &border_);

    /// <summary>
    /// Build and return an xml tree representing conditional_format_.
    /// </summary>
    static xml_node write_conditional_format(const conditional_format &conditional_format_);

    /// <summary>
    /// Build and return an xml tree representing fill_.
    /// </summary>
    static xml_node write_fill(const fill &fill_);

    /// <summary>
    /// Build and return an xml tree representing font_.
    /// </summary>
    static xml_node write_font(const font &font_);

    /// <summary>
    /// Build and return two xml trees, first=cell style and second=named style, representing named_style_.
    /// </summary>
    static std::pair<xml_node, xml_node> write_named_style(const named_style &named_style_);

    /// <summary>
    /// Build an xml tree representing all named styles in workbook into named_styles_node and cell_styles_node.
    /// Returns true on success.
    /// </summary>
    static std::pair<xml_node, xml_node> write_named_styles();

    /// <summary>
    /// Build and return an xml tree representing number_format_.
    /// </summary>
    static xml_node write_number_format(const number_format &number_format_);

    //
    // Non-static element readers (i.e. readers that modify internal state)
    //

    /// <summary>
    /// Read all borders from borders_node and add them to workbook.
    /// Return true on success.
    /// </summary>
    bool read_borders(const xml_node &borders_node);

    /// <summary>
    /// Read all fills from fills_node and add them to workbook.
    /// Return true on success.
    /// </summary>
    bool read_fills(const xml_node &fills_node);

    /// <summary>
    /// Read all fonts from fonts_node and add them to workbook.
    /// Return true on success.
    /// </summary>
    bool read_fonts(const xml_node &fonts_node);

    /// <summary>
    /// Read all colors from colors_node and add them to workbook.
    /// Return true on success.
    /// </summary>
    bool read_colors(const xml_node &colors_node);

    /// <summary>
    /// Read all named styles from named_style_node and cell_styles_node and add them to workbook.
    /// Return true on success.
    /// </summary>
    bool read_named_styles(const xml_node &named_styles_node);

    /// <summary>
    /// Read all borders from number_formats_node and add them to workbook.
    /// Return true on success.
    /// </summary>
    bool read_number_formats(const xml_node &number_formats_node);

    /// <summary>
    /// Read a single style from the given node. In styles.xml, this is an "xf" element.
    /// A style has an optional border id, fill id, font id, and number format id where
    /// each id is an index in the corresponding list of borders, etc. A style also has
    /// optional alignment and protection children. A style also defines whether each of
    /// these is "applied". For example, a style with a defined font id, font=# but with
    /// "applyFont=0" will not use the font in formatting.
    /// </summary>
    style read_style(const xml_node &style_node);

    //
    // Non-static element writers (i.e. writers that modify internal workbook)
    //

    /// <summary>
    /// Build an xml tree representing all borders in workbook into borders_node.
    /// Returns true on success.
    /// </summary>
    bool write_borders(xml_node &borders_node) const;

    /// <summary>
    /// Build an xml tree representing all conditional formats in workbook into conditional_formats_node.
    /// Returns true on success.
    /// </summary>
    bool write_conditional_formats(xml_node &conditional_formats_node) const;

    /// <summary>
    /// Build an xml tree representing all fills in workbook into fills_node.
    /// Returns true on success.
    /// </summary>
    bool write_fills(xml_node &fills_node) const;

    /// <summary>
    /// Build an xml tree representing all fonts in workbook into fonts_node.
    /// Returns true on success.
    /// </summary>
    bool write_fonts(xml_node &fonts_node) const;

    /// <summary>
    /// Build an xml tree representing all number formats in workbook into number_formats_node.
    /// Returns true on success.
    /// </summary>
    bool write_number_formats(xml_node fonts_node) const;

    /// <summary>
    /// Build an xml tree representing the given style into style_node.
    /// Returns true on success.
    /// </summary>
    bool write_style(const style &style_, xml_node &style_node) const;

  private:
    /// <summary>
    /// Set in the constructor, this workbook is used as the source or target for all writing or reading, respectively.
    /// </summary>
    workbook &workbook_;
};

} // namespace xlnt
