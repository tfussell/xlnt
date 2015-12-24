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
#include <xlnt/serialization/theme_serializer.hpp>
#include <xlnt/serialization/xml_document.hpp>
#include <xlnt/serialization/xml_node.hpp>

#include <detail/constants.hpp>

namespace xlnt {

// I have no idea what this stuff is. I hope it was worth it.
xml_document theme_serializer::write_theme(const theme &) const
{
    xml_document xml;

    auto theme_node = xml.add_child("a:theme");
    xml.add_namespace("a", constants::Namespace("drawingml"));
    theme_node.add_attribute("name", "Office Theme");

    auto theme_elements_node = theme_node.add_child("a:themeElements");
    auto clr_scheme_node = theme_elements_node.add_child("a:clrScheme");
    clr_scheme_node.add_attribute("name", "Office");

    struct scheme_element
    {
        std::string name;
        std::string sub_element_name;
        std::string val;
    };

    std::vector<scheme_element> scheme_elements = {
        { "a:dk1", "a:sysClr", "windowText" },  { "a:lt1", "a:sysClr", "window" },
        { "a:dk2", "a:srgbClr", "1F497D" },     { "a:lt2", "a:srgbClr", "EEECE1" },
        { "a:accent1", "a:srgbClr", "4F81BD" }, { "a:accent2", "a:srgbClr", "C0504D" },
        { "a:accent3", "a:srgbClr", "9BBB59" }, { "a:accent4", "a:srgbClr", "8064A2" },
        { "a:accent5", "a:srgbClr", "4BACC6" }, { "a:accent6", "a:srgbClr", "F79646" },
        { "a:hlink", "a:srgbClr", "0000FF" },   { "a:folHlink", "a:srgbClr", "800080" },
    };

    for (auto element : scheme_elements)
    {
        auto element_node = clr_scheme_node.add_child(element.name);
        element_node.add_child(element.sub_element_name).add_attribute("val", element.val);

        if (element.name == "a:dk1")
        {
            element_node.get_child(element.sub_element_name).add_attribute("lastClr", "000000");
        }
        else if (element.name == "a:lt1")
        {
            element_node.get_child(element.sub_element_name).add_attribute("lastClr", "FFFFFF");
        }
    }

    struct font_scheme
    {
        bool typeface;
        std::string script;
        std::string major;
        std::string minor;
    };

    std::vector<font_scheme> font_schemes = {
        { true, "a:latin", "Cambria", "Calibri" },
        { true, "a:ea", "", "" },
        { true, "a:cs", "", "" },
        { false, "Jpan", "\xef\xbc\xad\xef\xbc\xb3 \xef\xbc\xb0\xe3\x82\xb4\xe3\x82\xb7\xe3\x83\x83\xe3\x82\xaf",
          "\xef\xbc\xad\xef\xbc\xb3 \xef\xbc\xb0\xe3\x82\xb4\xe3\x82\xb7\xe3\x83\x83\xe3\x82\xaf" },
        { false, "Hang", "\xeb\xa7\x91\xec\x9d\x80 \xea\xb3\xa0\xeb\x94\x95",
          "\xeb\xa7\x91\xec\x9d\x80 \xea\xb3\xa0\xeb\x94\x95" },
        { false, "Hans", "\xe5\xae\x8b\xe4\xbd\x93", "\xe5\xae\x8b\xe4\xbd\x93" },
        { false, "Hant", "\xe6\x96\xb0\xe7\xb4\xb0\xe6\x98\x8e\xe9\xab\x94",
          "\xe6\x96\xb0\xe7\xb4\xb0\xe6\x98\x8e\xe9\xab\x94" },
        { false, "Arab", "Times New Roman", "Arial" },
        { false, "Hebr", "Times New Roman", "Arial" },
        { false, "Thai", "Tahoma", "Tahoma" },
        { false, "Ethi", "Nyala", "Nyala" },
        { false, "Beng", "Vrinda", "Vrinda" },
        { false, "Gujr", "Shruti", "Shruti" },
        { false, "Khmr", "MoolBoran", "DaunPenh" },
        { false, "Knda", "Tunga", "Tunga" },
        { false, "Guru", "Raavi", "Raavi" },
        { false, "Cans", "Euphemia", "Euphemia" },
        { false, "Cher", "Plantagenet Cherokee", "Plantagenet Cherokee" },
        { false, "Yiii", "Microsoft Yi Baiti", "Microsoft Yi Baiti" },
        { false, "Tibt", "Microsoft Himalaya", "Microsoft Himalaya" },
        { false, "Thaa", "MV Boli", "MV Boli" },
        { false, "Deva", "Mangal", "Mangal" },
        { false, "Telu", "Gautami", "Gautami" },
        { false, "Taml", "Latha", "Latha" },
        { false, "Syrc", "Estrangelo Edessa", "Estrangelo Edessa" },
        { false, "Orya", "Kalinga", "Kalinga" },
        { false, "Mlym", "Kartika", "Kartika" },
        { false, "Laoo", "DokChampa", "DokChampa" },
        { false, "Sinh", "Iskoola Pota", "Iskoola Pota" },
        { false, "Mong", "Mongolian Baiti", "Mongolian Baiti" },
        { false, "Viet", "Times New Roman", "Arial" },
        { false, "Uigh", "Microsoft Uighur", "Microsoft Uighur" }
    };

    auto font_scheme_node = theme_elements_node.add_child("a:fontScheme");
    font_scheme_node.add_attribute("name", "Office");

    auto major_fonts_node = font_scheme_node.add_child("a:majorFont");
    auto minor_fonts_node = font_scheme_node.add_child("a:minorFont");

    for (auto scheme : font_schemes)
    {
        if (scheme.typeface)
        {
            auto major_font_node = major_fonts_node.add_child(scheme.script);
            major_font_node.add_attribute("typeface", scheme.major);

            auto minor_font_node = minor_fonts_node.add_child(scheme.script);
            minor_font_node.add_attribute("typeface", scheme.minor);
        }
        else
        {
            auto major_font_node = major_fonts_node.add_child("a:font");
            major_font_node.add_attribute("script", scheme.script);
            major_font_node.add_attribute("typeface", scheme.major);

            auto minor_font_node = minor_fonts_node.add_child("a:font");
            minor_font_node.add_attribute("script", scheme.script);
            minor_font_node.add_attribute("typeface", scheme.minor);
        }
    }

    auto format_scheme_node = theme_elements_node.add_child("a:fmtScheme");
    format_scheme_node.add_attribute("name", "Office");

    auto fill_style_list_node = format_scheme_node.add_child("a:fillStyleLst");
    fill_style_list_node.add_child("a:solidFill").add_child("a:schemeClr").add_attribute("val", "phClr");

    auto grad_fill_node = fill_style_list_node.add_child("a:gradFill");
    grad_fill_node.add_attribute("rotWithShape", "1");

    auto grad_fill_list = grad_fill_node.add_child("a:gsLst");
    auto gs_node = grad_fill_list.add_child("a:gs");
    gs_node.add_attribute("pos", "0");
    auto scheme_color_node = gs_node.add_child("a:schemeClr");
    scheme_color_node.add_attribute("val", "phClr");
    scheme_color_node.add_child("a:tint").add_attribute("val", "50000");
    scheme_color_node.add_child("a:satMod").add_attribute("val", "300000");

    gs_node = grad_fill_list.add_child("a:gs");
    gs_node.add_attribute("pos", "35000");
    scheme_color_node = gs_node.add_child("a:schemeClr");
    scheme_color_node.add_attribute("val", "phClr");
    scheme_color_node.add_child("a:tint").add_attribute("val", "37000");
    scheme_color_node.add_child("a:satMod").add_attribute("val", "300000");

    gs_node = grad_fill_list.add_child("a:gs");
    gs_node.add_attribute("pos", "100000");
    scheme_color_node = gs_node.add_child("a:schemeClr");
    scheme_color_node.add_attribute("val", "phClr");
    scheme_color_node.add_child("a:tint").add_attribute("val", "15000");
    scheme_color_node.add_child("a:satMod").add_attribute("val", "350000");

    auto lin_node = grad_fill_node.add_child("a:lin");
    lin_node.add_attribute("ang", "16200000");
    lin_node.add_attribute("scaled", "1");

    grad_fill_node = fill_style_list_node.add_child("a:gradFill");
    grad_fill_node.add_attribute("rotWithShape", "1");

    grad_fill_list = grad_fill_node.add_child("a:gsLst");
    gs_node = grad_fill_list.add_child("a:gs");
    gs_node.add_attribute("pos", "0");
    scheme_color_node = gs_node.add_child("a:schemeClr");
    scheme_color_node.add_attribute("val", "phClr");
    scheme_color_node.add_child("a:shade").add_attribute("val", "51000");
    scheme_color_node.add_child("a:satMod").add_attribute("val", "130000");

    gs_node = grad_fill_list.add_child("a:gs");
    gs_node.add_attribute("pos", "80000");
    scheme_color_node = gs_node.add_child("a:schemeClr");
    scheme_color_node.add_attribute("val", "phClr");
    scheme_color_node.add_child("a:shade").add_attribute("val", "93000");
    scheme_color_node.add_child("a:satMod").add_attribute("val", "130000");

    gs_node = grad_fill_list.add_child("a:gs");
    gs_node.add_attribute("pos", "100000");
    scheme_color_node = gs_node.add_child("a:schemeClr");
    scheme_color_node.add_attribute("val", "phClr");
    scheme_color_node.add_child("a:shade").add_attribute("val", "94000");
    scheme_color_node.add_child("a:satMod").add_attribute("val", "135000");

    lin_node = grad_fill_node.add_child("a:lin");
    lin_node.add_attribute("ang", "16200000");
    lin_node.add_attribute("scaled", "0");

    auto line_style_list_node = format_scheme_node.add_child("a:lnStyleLst");

    auto ln_node = line_style_list_node.add_child("a:ln");
    ln_node.add_attribute("w", "9525");
    ln_node.add_attribute("cap", "flat");
    ln_node.add_attribute("cmpd", "sng");
    ln_node.add_attribute("algn", "ctr");

    auto solid_fill_node = ln_node.add_child("a:solidFill");
    scheme_color_node = solid_fill_node.add_child("a:schemeClr");
    scheme_color_node.add_attribute("val", "phClr");
    scheme_color_node.add_child("a:shade").add_attribute("val", "95000");
    scheme_color_node.add_child("a:satMod").add_attribute("val", "105000");
    ln_node.add_child("a:prstDash").add_attribute("val", "solid");

    ln_node = line_style_list_node.add_child("a:ln");
    ln_node.add_attribute("w", "25400");
    ln_node.add_attribute("cap", "flat");
    ln_node.add_attribute("cmpd", "sng");
    ln_node.add_attribute("algn", "ctr");

    solid_fill_node = ln_node.add_child("a:solidFill");
    scheme_color_node = solid_fill_node.add_child("a:schemeClr");
    scheme_color_node.add_attribute("val", "phClr");
    ln_node.add_child("a:prstDash").add_attribute("val", "solid");

    ln_node = line_style_list_node.add_child("a:ln");
    ln_node.add_attribute("w", "38100");
    ln_node.add_attribute("cap", "flat");
    ln_node.add_attribute("cmpd", "sng");
    ln_node.add_attribute("algn", "ctr");

    solid_fill_node = ln_node.add_child("a:solidFill");
    scheme_color_node = solid_fill_node.add_child("a:schemeClr");
    scheme_color_node.add_attribute("val", "phClr");
    ln_node.add_child("a:prstDash").add_attribute("val", "solid");

    auto effect_style_list_node = format_scheme_node.add_child("a:effectStyleLst");
    auto effect_style_node = effect_style_list_node.add_child("a:effectStyle");
    auto effect_list_node = effect_style_node.add_child("a:effectLst");
    auto outer_shadow_node = effect_list_node.add_child("a:outerShdw");
    outer_shadow_node.add_attribute("blurRad", "40000");
    outer_shadow_node.add_attribute("dist", "20000");
    outer_shadow_node.add_attribute("dir", "5400000");
    outer_shadow_node.add_attribute("rotWithShape", "0");
    auto srgb_clr_node = outer_shadow_node.add_child("a:srgbClr");
    srgb_clr_node.add_attribute("val", "000000");
    srgb_clr_node.add_child("a:alpha").add_attribute("val", "38000");

    effect_style_node = effect_style_list_node.add_child("a:effectStyle");
    effect_list_node = effect_style_node.add_child("a:effectLst");
    outer_shadow_node = effect_list_node.add_child("a:outerShdw");
    outer_shadow_node.add_attribute("blurRad", "40000");
    outer_shadow_node.add_attribute("dist", "23000");
    outer_shadow_node.add_attribute("dir", "5400000");
    outer_shadow_node.add_attribute("rotWithShape", "0");
    srgb_clr_node = outer_shadow_node.add_child("a:srgbClr");
    srgb_clr_node.add_attribute("val", "000000");
    srgb_clr_node.add_child("a:alpha").add_attribute("val", "35000");

    effect_style_node = effect_style_list_node.add_child("a:effectStyle");
    effect_list_node = effect_style_node.add_child("a:effectLst");
    outer_shadow_node = effect_list_node.add_child("a:outerShdw");
    outer_shadow_node.add_attribute("blurRad", "40000");
    outer_shadow_node.add_attribute("dist", "23000");
    outer_shadow_node.add_attribute("dir", "5400000");
    outer_shadow_node.add_attribute("rotWithShape", "0");
    srgb_clr_node = outer_shadow_node.add_child("a:srgbClr");
    srgb_clr_node.add_attribute("val", "000000");
    srgb_clr_node.add_child("a:alpha").add_attribute("val", "35000");
    auto scene3d_node = effect_style_node.add_child("a:scene3d");
    auto camera_node = scene3d_node.add_child("a:camera");
    camera_node.add_attribute("prst", "orthographicFront");
    auto rot_node = camera_node.add_child("a:rot");
    rot_node.add_attribute("lat", "0");
    rot_node.add_attribute("lon", "0");
    rot_node.add_attribute("rev", "0");
    auto light_rig_node = scene3d_node.add_child("a:lightRig");
    light_rig_node.add_attribute("rig", "threePt");
    light_rig_node.add_attribute("dir", "t");
    rot_node = light_rig_node.add_child("a:rot");
    rot_node.add_attribute("lat", "0");
    rot_node.add_attribute("lon", "0");
    rot_node.add_attribute("rev", "1200000");

    auto bevel_node = effect_style_node.add_child("a:sp3d").add_child("a:bevelT");
    bevel_node.add_attribute("w", "63500");
    bevel_node.add_attribute("h", "25400");

    auto bg_fill_style_list_node = format_scheme_node.add_child("a:bgFillStyleLst");

    bg_fill_style_list_node.add_child("a:solidFill").add_child("a:schemeClr").add_attribute("val", "phClr");

    grad_fill_node = bg_fill_style_list_node.add_child("a:gradFill");
    grad_fill_node.add_attribute("rotWithShape", "1");

    grad_fill_list = grad_fill_node.add_child("a:gsLst");
    gs_node = grad_fill_list.add_child("a:gs");
    gs_node.add_attribute("pos", "0");
    scheme_color_node = gs_node.add_child("a:schemeClr");
    scheme_color_node.add_attribute("val", "phClr");
    scheme_color_node.add_child("a:tint").add_attribute("val", "40000");
    scheme_color_node.add_child("a:satMod").add_attribute("val", "350000");

    gs_node = grad_fill_list.add_child("a:gs");
    gs_node.add_attribute("pos", "40000");
    scheme_color_node = gs_node.add_child("a:schemeClr");
    scheme_color_node.add_attribute("val", "phClr");
    scheme_color_node.add_child("a:tint").add_attribute("val", "45000");
    scheme_color_node.add_child("a:shade").add_attribute("val", "99000");
    scheme_color_node.add_child("a:satMod").add_attribute("val", "350000");

    gs_node = grad_fill_list.add_child("a:gs");
    gs_node.add_attribute("pos", "100000");
    scheme_color_node = gs_node.add_child("a:schemeClr");
    scheme_color_node.add_attribute("val", "phClr");
    scheme_color_node.add_child("a:shade").add_attribute("val", "20000");
    scheme_color_node.add_child("a:satMod").add_attribute("val", "255000");

    auto path_node = grad_fill_node.add_child("a:path");
    path_node.add_attribute("path", "circle");
    auto fill_to_rect_node = path_node.add_child("a:fillToRect");
    fill_to_rect_node.add_attribute("l", "50000");
    fill_to_rect_node.add_attribute("t", "-80000");
    fill_to_rect_node.add_attribute("r", "50000");
    fill_to_rect_node.add_attribute("b", "180000");

    grad_fill_node = bg_fill_style_list_node.add_child("a:gradFill");
    grad_fill_node.add_attribute("rotWithShape", "1");

    grad_fill_list = grad_fill_node.add_child("a:gsLst");
    gs_node = grad_fill_list.add_child("a:gs");
    gs_node.add_attribute("pos", "0");
    scheme_color_node = gs_node.add_child("a:schemeClr");
    scheme_color_node.add_attribute("val", "phClr");
    scheme_color_node.add_child("a:tint").add_attribute("val", "80000");
    scheme_color_node.add_child("a:satMod").add_attribute("val", "300000");

    gs_node = grad_fill_list.add_child("a:gs");
    gs_node.add_attribute("pos", "100000");
    scheme_color_node = gs_node.add_child("a:schemeClr");
    scheme_color_node.add_attribute("val", "phClr");
    scheme_color_node.add_child("a:shade").add_attribute("val", "30000");
    scheme_color_node.add_child("a:satMod").add_attribute("val", "200000");

    path_node = grad_fill_node.add_child("a:path");
    path_node.add_attribute("path", "circle");
    fill_to_rect_node = path_node.add_child("a:fillToRect");
    fill_to_rect_node.add_attribute("l", "50000");
    fill_to_rect_node.add_attribute("t", "50000");
    fill_to_rect_node.add_attribute("r", "50000");
    fill_to_rect_node.add_attribute("b", "50000");

    theme_node.add_child("a:objectDefaults");
    theme_node.add_child("a:extraClrSchemeLst");

    return xml;
}

} // namespace xlnt
