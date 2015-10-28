// Copyright (c) 2015 Thomas Fussell
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

#include "xlnt/workbook/workbook.hpp"

namespace pugi {
class xml_node;
} // namespace pugi

namespace xlnt {

class border;
class fill;
class font;
class named_style;
class number_format;
class style;
class zip_file;
    
class style_reader
{
public:
    style_reader(workbook &wb);
    
    void read_styles(zip_file &archive);
    
    const std::vector<style> &get_styles() const { return styles_; }
    
    const std::vector<border> &get_borders() const { return borders_; }
    const std::vector<fill> &get_fills() const { return fills_; }
    const std::vector<font> &get_fonts() const { return fonts_; }
    const std::vector<number_format> &get_number_formats() const { return number_formats_; }
    const std::vector<std::string> &get_colors() const { return colors_; }
    
private:
    style read_style(pugi::xml_node stylesheet_node, pugi::xml_node xf_node);
    
    void read_borders(pugi::xml_node borders_node);
    void read_fills(pugi::xml_node fills_node);
    void read_fonts(pugi::xml_node fonts_node);
    void read_number_formats(pugi::xml_node num_fmt_node);
    void read_colors(pugi::xml_node colors_node);
    
    void read_dxfs();
    void read_named_styles(pugi::xml_node named_styles_node);
    void read_style_names();
    void read_cell_styles();
    void read_xfs();
    void read_style_table();

    workbook wb_;
    
    std::vector<style> styles_;
    
    std::vector<std::string> colors_;
    std::vector<border> borders_;
    std::vector<fill> fills_;
    std::vector<font> fonts_;
    std::vector<number_format> number_formats_;
};

} // namespace xlnt
