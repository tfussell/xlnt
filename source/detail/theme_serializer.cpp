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

#include <pugixml.hpp>
#include <vector>

#include <detail/constants.hpp>
#include <detail/theme_serializer.hpp>

namespace xlnt {

// I have no idea what this stuff is. I hope it was worth it.
void theme_serializer::write_theme(const theme &, pugi::xml_document &xml) const
{
    auto theme_node = xml.append_child("a:theme");
    theme_node.append_attribute("xmlns:a").set_value(constants::get_namespace("drawingml").c_str());
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
        auto element_node = clr_scheme_node.append_child(element.name.c_str());
        element_node.append_child(element.sub_element_name.c_str()).append_attribute("val").set_value(element.val.c_str());

        if (element.name == "a:dk1")
        {
            element_node.child(element.sub_element_name.c_str()).append_attribute("lastClr").set_value("000000");
        }
        else if (element.name == "a:lt1")
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

    auto font_scheme_node = theme_elements_node.append_child("a:fontScheme");
    font_scheme_node.append_attribute("name").set_value("Office");

    auto major_fonts_node = font_scheme_node.append_child("a:majorFont");
    auto minor_fonts_node = font_scheme_node.append_child("a:minorFont");

    for (auto scheme : font_schemes)
    {
        if (scheme.typeface)
        {
            auto major_font_node = major_fonts_node.append_child(scheme.script.c_str());
            major_font_node.append_attribute("typeface").set_value(scheme.major.c_str());

            auto minor_font_node = minor_fonts_node.append_child(scheme.script.c_str());
            minor_font_node.append_attribute("typeface").set_value(scheme.minor.c_str());
        }
        else
        {
            auto major_font_node = major_fonts_node.append_child("a:font");
            major_font_node.append_attribute("script").set_value(scheme.script.c_str());
            major_font_node.append_attribute("typeface").set_value(scheme.major.c_str());

            auto minor_font_node = minor_fonts_node.append_child("a:font");
            minor_font_node.append_attribute("script").set_value(scheme.script.c_str());
            minor_font_node.append_attribute("typeface").set_value(scheme.minor.c_str());
        }
    }

    auto format_scheme_node = theme_elements_node.append_child("a:fmtScheme");
    format_scheme_node.append_attribute("name").set_value("Office");

    auto fill_style_list_node = format_scheme_node.append_child("a:fillStyleLst");
    fill_style_list_node.append_child("a:solidFill").append_child("a:schemeClr").append_attribute("val").set_value("phClr");

    auto grad_fill_node = fill_style_list_node.append_child("a:gradFill");
    grad_fill_node.append_attribute("rotWithShape").set_value("1");

    auto grad_fill_list = grad_fill_node.append_child("a:gsLst");
    auto gs_node = grad_fill_list.append_child("a:gs");
    gs_node.append_attribute("pos").set_value("0");
    auto scheme_color_node = gs_node.append_child("a:schemeClr");
    scheme_color_node.append_attribute("val").set_value("phClr");
    scheme_color_node.append_child("a:tint").append_attribute("val").set_value("50000");
    scheme_color_node.append_child("a:satMod").append_attribute("val").set_value("300000");

    gs_node = grad_fill_list.append_child("a:gs");
    gs_node.append_attribute("pos").set_value("35000");
    scheme_color_node = gs_node.append_child("a:schemeClr");
    scheme_color_node.append_attribute("val").set_value("phClr");
    scheme_color_node.append_child("a:tint").append_attribute("val").set_value("37000");
    scheme_color_node.append_child("a:satMod").append_attribute("val").set_value("300000");

    gs_node = grad_fill_list.append_child("a:gs");
    gs_node.append_attribute("pos").set_value("100000");
    scheme_color_node = gs_node.append_child("a:schemeClr");
    scheme_color_node.append_attribute("val").set_value("phClr");
    scheme_color_node.append_child("a:tint").append_attribute("val").set_value("15000");
    scheme_color_node.append_child("a:satMod").append_attribute("val").set_value("350000");

    auto lin_node = grad_fill_node.append_child("a:lin");
    lin_node.append_attribute("ang").set_value("16200000");
    lin_node.append_attribute("scaled").set_value("1");

    grad_fill_node = fill_style_list_node.append_child("a:gradFill");
    grad_fill_node.append_attribute("rotWithShape").set_value("1");

    grad_fill_list = grad_fill_node.append_child("a:gsLst");
    gs_node = grad_fill_list.append_child("a:gs");
    gs_node.append_attribute("pos").set_value("0");
    scheme_color_node = gs_node.append_child("a:schemeClr");
    scheme_color_node.append_attribute("val").set_value("phClr");
    scheme_color_node.append_child("a:shade").append_attribute("val").set_value("51000");
    scheme_color_node.append_child("a:satMod").append_attribute("val").set_value("130000");

    gs_node = grad_fill_list.append_child("a:gs");
    gs_node.append_attribute("pos").set_value("80000");
    scheme_color_node = gs_node.append_child("a:schemeClr");
    scheme_color_node.append_attribute("val").set_value("phClr");
    scheme_color_node.append_child("a:shade").append_attribute("val").set_value("93000");
    scheme_color_node.append_child("a:satMod").append_attribute("val").set_value("130000");

    gs_node = grad_fill_list.append_child("a:gs");
    gs_node.append_attribute("pos").set_value("100000");
    scheme_color_node = gs_node.append_child("a:schemeClr");
    scheme_color_node.append_attribute("val").set_value("phClr");
    scheme_color_node.append_child("a:shade").append_attribute("val").set_value("94000");
    scheme_color_node.append_child("a:satMod").append_attribute("val").set_value("135000");

    lin_node = grad_fill_node.append_child("a:lin");
    lin_node.append_attribute("ang").set_value("16200000");
    lin_node.append_attribute("scaled").set_value("0");

    auto line_style_list_node = format_scheme_node.append_child("a:lnStyleLst");

    auto ln_node = line_style_list_node.append_child("a:ln");
    ln_node.append_attribute("w").set_value("9525");
    ln_node.append_attribute("cap").set_value("flat");
    ln_node.append_attribute("cmpd").set_value("sng");
    ln_node.append_attribute("algn").set_value("ctr");

    auto solid_fill_node = ln_node.append_child("a:solidFill");
    scheme_color_node = solid_fill_node.append_child("a:schemeClr");
    scheme_color_node.append_attribute("val").set_value("phClr");
    scheme_color_node.append_child("a:shade").append_attribute("val").set_value("95000");
    scheme_color_node.append_child("a:satMod").append_attribute("val").set_value("105000");
    ln_node.append_child("a:prstDash").append_attribute("val").set_value("solid");

    ln_node = line_style_list_node.append_child("a:ln");
    ln_node.append_attribute("w").set_value("25400");
    ln_node.append_attribute("cap").set_value("flat");
    ln_node.append_attribute("cmpd").set_value("sng");
    ln_node.append_attribute("algn").set_value("ctr");

    solid_fill_node = ln_node.append_child("a:solidFill");
    scheme_color_node = solid_fill_node.append_child("a:schemeClr");
    scheme_color_node.append_attribute("val").set_value("phClr");
    ln_node.append_child("a:prstDash").append_attribute("val").set_value("solid");

    ln_node = line_style_list_node.append_child("a:ln");
    ln_node.append_attribute("w").set_value("38100");
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
    outer_shadow_node.append_attribute("blurRad").set_value("40000");
    outer_shadow_node.append_attribute("dist").set_value("20000");
    outer_shadow_node.append_attribute("dir").set_value("5400000");
    outer_shadow_node.append_attribute("rotWithShape").set_value("0");
    auto srgb_clr_node = outer_shadow_node.append_child("a:srgbClr");
    srgb_clr_node.append_attribute("val").set_value("000000");
    srgb_clr_node.append_child("a:alpha").append_attribute("val").set_value("38000");

    effect_style_node = effect_style_list_node.append_child("a:effectStyle");
    effect_list_node = effect_style_node.append_child("a:effectLst");
    outer_shadow_node = effect_list_node.append_child("a:outerShdw");
    outer_shadow_node.append_attribute("blurRad").set_value("40000");
    outer_shadow_node.append_attribute("dist").set_value("23000");
    outer_shadow_node.append_attribute("dir").set_value("5400000");
    outer_shadow_node.append_attribute("rotWithShape").set_value("0");
    srgb_clr_node = outer_shadow_node.append_child("a:srgbClr");
    srgb_clr_node.append_attribute("val").set_value("000000");
    srgb_clr_node.append_child("a:alpha").append_attribute("val").set_value("35000");

    effect_style_node = effect_style_list_node.append_child("a:effectStyle");
    effect_list_node = effect_style_node.append_child("a:effectLst");
    outer_shadow_node = effect_list_node.append_child("a:outerShdw");
    outer_shadow_node.append_attribute("blurRad").set_value("40000");
    outer_shadow_node.append_attribute("dist").set_value("23000");
    outer_shadow_node.append_attribute("dir").set_value("5400000");
    outer_shadow_node.append_attribute("rotWithShape").set_value("0");
    srgb_clr_node = outer_shadow_node.append_child("a:srgbClr");
    srgb_clr_node.append_attribute("val").set_value("000000");
    srgb_clr_node.append_child("a:alpha").append_attribute("val").set_value("35000");
    auto scene3d_node = effect_style_node.append_child("a:scene3d");
    auto camera_node = scene3d_node.append_child("a:camera");
    camera_node.append_attribute("prst").set_value("orthographicFront");
    auto rot_node = camera_node.append_child("a:rot");
    rot_node.append_attribute("lat").set_value("0");
    rot_node.append_attribute("lon").set_value("0");
    rot_node.append_attribute("rev").set_value("0");
    auto light_rig_node = scene3d_node.append_child("a:lightRig");
    light_rig_node.append_attribute("rig").set_value("threePt");
    light_rig_node.append_attribute("dir").set_value("t");
    rot_node = light_rig_node.append_child("a:rot");
    rot_node.append_attribute("lat").set_value("0");
    rot_node.append_attribute("lon").set_value("0");
    rot_node.append_attribute("rev").set_value("1200000");

    auto bevel_node = effect_style_node.append_child("a:sp3d").append_child("a:bevelT");
    bevel_node.append_attribute("w").set_value("63500");
    bevel_node.append_attribute("h").set_value("25400");

    auto bg_fill_style_list_node = format_scheme_node.append_child("a:bgFillStyleLst");

    bg_fill_style_list_node.append_child("a:solidFill").append_child("a:schemeClr").append_attribute("val").set_value("phClr");

    grad_fill_node = bg_fill_style_list_node.append_child("a:gradFill");
    grad_fill_node.append_attribute("rotWithShape").set_value("1");

    grad_fill_list = grad_fill_node.append_child("a:gsLst");
    gs_node = grad_fill_list.append_child("a:gs");
    gs_node.append_attribute("pos").set_value("0");
    scheme_color_node = gs_node.append_child("a:schemeClr");
    scheme_color_node.append_attribute("val").set_value("phClr");
    scheme_color_node.append_child("a:tint").append_attribute("val").set_value("40000");
    scheme_color_node.append_child("a:satMod").append_attribute("val").set_value("350000");

    gs_node = grad_fill_list.append_child("a:gs");
    gs_node.append_attribute("pos").set_value("40000");
    scheme_color_node = gs_node.append_child("a:schemeClr");
    scheme_color_node.append_attribute("val").set_value("phClr");
    scheme_color_node.append_child("a:tint").append_attribute("val").set_value("45000");
    scheme_color_node.append_child("a:shade").append_attribute("val").set_value("99000");
    scheme_color_node.append_child("a:satMod").append_attribute("val").set_value("350000");

    gs_node = grad_fill_list.append_child("a:gs");
    gs_node.append_attribute("pos").set_value("100000");
    scheme_color_node = gs_node.append_child("a:schemeClr");
    scheme_color_node.append_attribute("val").set_value("phClr");
    scheme_color_node.append_child("a:shade").append_attribute("val").set_value("20000");
    scheme_color_node.append_child("a:satMod").append_attribute("val").set_value("255000");

    auto path_node = grad_fill_node.append_child("a:path");
    path_node.append_attribute("path").set_value("circle");
    auto fill_to_rect_node = path_node.append_child("a:fillToRect");
    fill_to_rect_node.append_attribute("l").set_value("50000");
    fill_to_rect_node.append_attribute("t").set_value("-80000");
    fill_to_rect_node.append_attribute("r").set_value("50000");
    fill_to_rect_node.append_attribute("b").set_value("180000");

    grad_fill_node = bg_fill_style_list_node.append_child("a:gradFill");
    grad_fill_node.append_attribute("rotWithShape").set_value("1");

    grad_fill_list = grad_fill_node.append_child("a:gsLst");
    gs_node = grad_fill_list.append_child("a:gs");
    gs_node.append_attribute("pos").set_value("0");
    scheme_color_node = gs_node.append_child("a:schemeClr");
    scheme_color_node.append_attribute("val").set_value("phClr");
    scheme_color_node.append_child("a:tint").append_attribute("val").set_value("80000");
    scheme_color_node.append_child("a:satMod").append_attribute("val").set_value("300000");

    gs_node = grad_fill_list.append_child("a:gs");
    gs_node.append_attribute("pos").set_value("100000");
    scheme_color_node = gs_node.append_child("a:schemeClr");
    scheme_color_node.append_attribute("val").set_value("phClr");
    scheme_color_node.append_child("a:shade").append_attribute("val").set_value("30000");
    scheme_color_node.append_child("a:satMod").append_attribute("val").set_value("200000");

    path_node = grad_fill_node.append_child("a:path");
    path_node.append_attribute("path").set_value("circle");
    fill_to_rect_node = path_node.append_child("a:fillToRect");
    fill_to_rect_node.append_attribute("l").set_value("50000");
    fill_to_rect_node.append_attribute("t").set_value("50000");
    fill_to_rect_node.append_attribute("r").set_value("50000");
    fill_to_rect_node.append_attribute("b").set_value("50000");

    theme_node.append_child("a:objectDefaults");
    theme_node.append_child("a:extraClrSchemeLst");
}

} // namespace xlnt
