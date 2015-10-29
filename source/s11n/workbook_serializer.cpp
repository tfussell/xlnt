#include <xlnt/s11n/workbook_serializer.hpp>
#include <xlnt/common/datetime.hpp>
#include <xlnt/common/relationship.hpp>
#include <xlnt/s11n/xml_document.hpp>
#include <xlnt/s11n/xml_node.hpp>
#include <xlnt/workbook/document_properties.hpp>
#include <xlnt/workbook/manifest.hpp>
#include <xlnt/workbook/workbook.hpp>

#include "detail/constants.hpp"

namespace {
    
xlnt::datetime w3cdtf_to_datetime(const std::string &string)
{
    xlnt::datetime result(1900, 1, 1);
    auto separator_index = string.find('-');
    result.year = std::stoi(string.substr(0, separator_index));
    result.month = std::stoi(string.substr(separator_index + 1, string.find('-', separator_index + 1)));
    separator_index = string.find('-', separator_index + 1);
    result.day = std::stoi(string.substr(separator_index + 1, string.find('T', separator_index + 1)));
    separator_index = string.find('T', separator_index + 1);
    result.hour = std::stoi(string.substr(separator_index + 1, string.find(':', separator_index + 1)));
    separator_index = string.find(':', separator_index + 1);
    result.minute = std::stoi(string.substr(separator_index + 1, string.find(':', separator_index + 1)));
    separator_index = string.find(':', separator_index + 1);
    result.second = std::stoi(string.substr(separator_index + 1, string.find('Z', separator_index + 1)));
    return result;
}
    
std::string fill(const std::string &string, std::size_t length = 2)
{
    if(string.size() >= length)
    {
        return string;
    }
    
    return std::string(length - string.size(), '0') + string;
}

std::string datetime_to_w3cdtf(const xlnt::datetime &dt)
{
    return std::to_string(dt.year) + "-" + fill(std::to_string(dt.month)) + "-" + fill(std::to_string(dt.day)) + "T" + fill(std::to_string(dt.hour)) + ":" + fill(std::to_string(dt.minute)) + ":" + fill(std::to_string(dt.second)) + "Z";
}

} // namespace

namespace xlnt {
    
    /*
std::vector<std::pair<std::string, std::string>> workbook_serializer::read_sheets(zip_file &archive)
{
    std::string ns;
    
    for(auto child : doc.children())
    {
        std::string name = child.name();
        
        if(name.find(':') != std::string::npos)
        {
            auto colon_index = name.find(':');
            ns = name.substr(0, colon_index);
            break;
        }
    }
    
    auto with_ns = [&](const std::string &base) { return ns.empty() ? base : ns + ":" + base; };
    
    auto root_node = doc.child(with_ns("workbook").c_str());
    auto sheets_node = root_node.child(with_ns("sheets").c_str());
    
    std::vector<std::pair<std::string, std::string>> sheets;
    
    // store temp because pugixml iteration uses the internal char array multiple times
    auto sheet_element_name = with_ns("sheet");
    
    for(auto sheet_node : sheets_node.children(sheet_element_name.c_str()))
    {
        std::string id = sheet_node.attribute("r:id").as_string();
        std::string name = sheet_node.attribute("name").as_string();
        sheets.push_back(std::make_pair(id, name));
    }
    
    return sheets;
}
     */

void workbook_serializer::read_properties_core(const xml_document &xml)
{
    auto &props = wb_.get_properties();
    auto root_node = xml.root();
    
    props.excel_base_date = calendar::windows_1900;
    
    if(root_node.has_child("dc:creator"))
    {
        props.creator = root_node.get_child("dc:creator").get_text();
    }
    if(root_node.has_child("cp:lastModifiedBy"))
    {
        props.last_modified_by = root_node.get_child("cp:lastModifiedBy").get_text();
    }
    if(root_node.has_child("dcterms:created"))
    {
        std::string created_string = root_node.get_child("dcterms:created").get_text();
        props.created = w3cdtf_to_datetime(created_string);
    }
    if(root_node.has_child("dcterms:modified"))
    {
        std::string modified_string = root_node.get_child("dcterms:modified").get_text();
        props.modified = w3cdtf_to_datetime(modified_string);
    }
}

std::string workbook_serializer::read_dimension(const std::string &xml_string)
{
    pugi::xml_document doc;
    doc.load(xml_string.c_str());
    auto root_node = doc.child("worksheet");
    auto dimension_node = root_node.child("dimension");
    std::string dimension = dimension_node.attribute("ref").as_string();
    return dimension;
}

std::vector<relationship> workbook_serializer::read_relationships(zip_file &archive, const std::string &filename)
{
    auto filename_separator_index = filename.find_last_of('/');
    auto basename = filename.substr(filename_separator_index + 1);
    auto dirname = filename.substr(0, filename_separator_index);
    auto rels_filename = dirname + "/_rels/" + basename + ".rels";
    
    pugi::xml_document doc;
    auto content = archive.read(rels_filename);
    doc.load(content.c_str());
    
    auto root_node = doc.child("Relationships");
    
    std::vector<relationship> relationships;
    
    for(auto relationship : root_node.children("Relationship"))
    {
        std::string id = relationship.attribute("Id").as_string();
        std::string type = relationship.attribute("Type").as_string();
        std::string target = relationship.attribute("Target").as_string();
        
        if(target[0] != '/' && target.substr(0, 2) != "..")
        {
            target = dirname + "/" + target;
        }
        
        if(target[0] == '/')
        {
            target = target.substr(1);
        }
        
        relationships.push_back(xlnt::relationship(type, id, target));
    }
    
    return relationships;
}

std::string workbook_serializer::determine_document_type(const manifest &manifest)
{
    if(!manifest.has_override_type(constants::ArcWorkbook))
    {
        return "unsupported";
    }
    
    std::string type = manifest.get_override_type(constants::ArcWorkbook);
    
    if(type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
    {
        return "excel";
    }
    else if(type == "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml")
    {
        return "powerpoint";
    }
    else if(type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")
    {
        return "word";
    }
    
    return "unsupported";
}

std::vector<std::pair<std::string, std::string>> workbook_serializer::detect_worksheets(zip_file &archive)
{
    static const std::string ValidWorksheet = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
    
    auto content_types = read_content_types(archive);
    std::vector<std::string> valid_sheets;
    
    for(const auto &content_type : content_types)
    {
        if(content_type.second == ValidWorksheet)
        {
            valid_sheets.push_back(content_type.first);
        }
    }
    
    auto workbook_relationships = read_relationships(archive, "xl/workbook.xml");
    std::vector<std::pair<std::string, std::string>> result;
    
    for(const auto &ws : read_sheets(archive))
    {
        auto rel = *std::find_if(workbook_relationships.begin(), workbook_relationships.end(), [&](const relationship &r) { return r.get_id() == ws.first; });
        auto target = rel.get_target_uri();
        
        if(std::find(valid_sheets.begin(), valid_sheets.end(), "/" + target) != valid_sheets.end())
        {
            result.push_back({target, ws.second});
        }
    }
    
    return result;
}

std::string workbook_serializer::write_shared_strings(const std::vector<std::string> &string_table)
{
    pugi::xml_document doc;
    auto root_node = doc.append_child("sst");
    root_node.append_attribute("xmlns").set_value("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    root_node.append_attribute("uniqueCount").set_value((int)string_table.size());
    
    for(auto string : string_table)
    {
        root_node.append_child("si").append_child("t").text().set(string.c_str());
    }
    
    std::stringstream ss;
    doc.save(ss);
    
    return ss.str();
}

std::string workbook_serializer::write_properties_core(const document_properties &prop)
{
    pugi::xml_document doc;
    auto root_node = doc.append_child("cp:coreProperties");
    root_node.append_attribute("xmlns:cp").set_value("http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
    root_node.append_attribute("xmlns:dc").set_value("http://purl.org/dc/elements/1.1/");
    root_node.append_attribute("xmlns:dcmitype").set_value("http://purl.org/dc/dcmitype/");
    root_node.append_attribute("xmlns:dcterms").set_value("http://purl.org/dc/terms/");
    root_node.append_attribute("xmlns:xsi").set_value("http://www.w3.org/2001/XMLSchema-instance");
    
    root_node.append_child("dc:creator").text().set(prop.creator.c_str());
    root_node.append_child("cp:lastModifiedBy").text().set(prop.last_modified_by.c_str());
    root_node.append_child("dcterms:created").text().set(datetime_to_w3cdtf(prop.created).c_str());
    root_node.child("dcterms:created").append_attribute("xsi:type").set_value("dcterms:W3CDTF");
    root_node.append_child("dcterms:modified").text().set(datetime_to_w3cdtf(prop.modified).c_str());
    root_node.child("dcterms:modified").append_attribute("xsi:type").set_value("dcterms:W3CDTF");
    root_node.append_child("dc:title").text().set(prop.title.c_str());
    root_node.append_child("dc:description");
    root_node.append_child("dc:subject");
    root_node.append_child("cp:keywords");
    root_node.append_child("cp:category");
    
    std::stringstream ss;
    doc.save(ss);
    
    return ss.str();
}

std::string workbook_serializer::write_worksheet_rels(worksheet ws)
{
    return write_relationships(ws.get_relationships(), "");
}

std::string workbook_serializer::write_theme()
{
    pugi::xml_document doc;
    auto theme_node = doc.append_child("a:theme");
    theme_node.append_attribute("xmlns:a").set_value(constants::Namespaces.at("drawingml").c_str());
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
    
    std::vector<scheme_element> scheme_elements =
    {
        {"a:dk1", "a:sysClr", "windowText"},
        {"a:lt1", "a:sysClr", "window"},
        {"a:dk2", "a:srgbClr", "1F497D"},
        {"a:lt2", "a:srgbClr", "EEECE1"},
        {"a:accent1", "a:srgbClr", "4F81BD"},
        {"a:accent2", "a:srgbClr", "C0504D"},
        {"a:accent3", "a:srgbClr", "9BBB59"},
        {"a:accent4", "a:srgbClr", "8064A2"},
        {"a:accent5", "a:srgbClr", "4BACC6"},
        {"a:accent6", "a:srgbClr", "F79646"},
        {"a:hlink", "a:srgbClr", "0000FF"},
        {"a:folHlink", "a:srgbClr", "800080"},
    };
    
    for(auto element : scheme_elements)
    {
        auto element_node = clr_scheme_node.append_child(element.name.c_str());
        element_node.append_child(element.sub_element_name.c_str()).append_attribute("val").set_value(element.val.c_str());
        
        if(element.name == "a:dk1")
        {
            element_node.child(element.sub_element_name.c_str()).append_attribute("lastClr").set_value("000000");
        }
        else if(element.name == "a:lt1")
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
    
    std::vector<font_scheme> font_schemes =
    {
        {true, "a:latin", "Cambria", "Calibri"},
        {true, "a:ea", "", ""},
        {true, "a:cs", "", ""},
        {false, "Jpan", "\xef\xbc\xad\xef\xbc\xb3 \xef\xbc\xb0\xe3\x82\xb4\xe3\x82\xb7\xe3\x83\x83\xe3\x82\xaf", "\xef\xbc\xad\xef\xbc\xb3 \xef\xbc\xb0\xe3\x82\xb4\xe3\x82\xb7\xe3\x83\x83\xe3\x82\xaf"},
        {false, "Hang", "\xeb\xa7\x91\xec\x9d\x80 \xea\xb3\xa0\xeb\x94\x95", "\xeb\xa7\x91\xec\x9d\x80 \xea\xb3\xa0\xeb\x94\x95"},
        {false, "Hans", "\xe5\xae\x8b\xe4\xbd\x93", "\xe5\xae\x8b\xe4\xbd\x93"},
        {false, "Hant", "\xe6\x96\xb0\xe7\xb4\xb0\xe6\x98\x8e\xe9\xab\x94", "\xe6\x96\xb0\xe7\xb4\xb0\xe6\x98\x8e\xe9\xab\x94"},
        {false, "Arab", "Times New Roman", "Arial"},
        {false, "Hebr", "Times New Roman", "Arial"},
        {false, "Thai", "Tahoma", "Tahoma"},
        {false, "Ethi", "Nyala", "Nyala"},
        {false, "Beng", "Vrinda", "Vrinda"},
        {false, "Gujr", "Shruti", "Shruti"},
        {false, "Khmr", "MoolBoran", "DaunPenh"},
        {false, "Knda", "Tunga", "Tunga"},
        {false, "Guru", "Raavi", "Raavi"},
        {false, "Cans", "Euphemia", "Euphemia"},
        {false, "Cher", "Plantagenet Cherokee", "Plantagenet Cherokee"},
        {false, "Yiii", "Microsoft Yi Baiti", "Microsoft Yi Baiti"},
        {false, "Tibt", "Microsoft Himalaya", "Microsoft Himalaya"},
        {false, "Thaa", "MV Boli", "MV Boli"},
        {false, "Deva", "Mangal", "Mangal"},
        {false, "Telu", "Gautami", "Gautami"},
        {false, "Taml", "Latha", "Latha"},
        {false, "Syrc", "Estrangelo Edessa", "Estrangelo Edessa"},
        {false, "Orya", "Kalinga", "Kalinga"},
        {false, "Mlym", "Kartika", "Kartika"},
        {false, "Laoo", "DokChampa", "DokChampa"},
        {false, "Sinh", "Iskoola Pota", "Iskoola Pota"},
        {false, "Mong", "Mongolian Baiti", "Mongolian Baiti"},
        {false, "Viet", "Times New Roman", "Arial"},
        {false, "Uigh", "Microsoft Uighur", "Microsoft Uighur"}
    };
    
    auto font_scheme_node = theme_elements_node.append_child("a:fontScheme");
    font_scheme_node.append_attribute("name").set_value("Office");
    
    auto major_fonts_node = font_scheme_node.append_child("a:majorFont");
    auto minor_fonts_node = font_scheme_node.append_child("a:minorFont");
    
    for(auto scheme : font_schemes)
    {
        pugi::xml_node major_font_node, minor_font_node;
        
        if(scheme.typeface)
        {
            major_font_node = major_fonts_node.append_child(scheme.script.c_str());
            minor_font_node = minor_fonts_node.append_child(scheme.script.c_str());
        }
        else
        {
            major_font_node = major_fonts_node.append_child("a:font");
            major_font_node.append_attribute("script").set_value(scheme.script.c_str());
            minor_font_node = minor_fonts_node.append_child("a:font");
            minor_font_node.append_attribute("script").set_value(scheme.script.c_str());
        }
        
        major_font_node.append_attribute("typeface").set_value(scheme.major.c_str());
        minor_font_node.append_attribute("typeface").set_value(scheme.minor.c_str());
    }
    
    auto format_scheme_node = theme_elements_node.append_child("a:fmtScheme");
    format_scheme_node.append_attribute("name").set_value("Office");
    
    auto fill_style_list_node = format_scheme_node.append_child("a:fillStyleLst");
    fill_style_list_node.append_child("a:solidFill").append_child("a:schemeClr").append_attribute("val").set_value("phClr");
    
    auto grad_fill_node = fill_style_list_node.append_child("a:gradFill");
    grad_fill_node.append_attribute("rotWithShape").set_value(1);
    
    auto grad_fill_list = grad_fill_node.append_child("a:gsLst");
    auto gs_node = grad_fill_list.append_child("a:gs");
    gs_node.append_attribute("pos").set_value(0);
    auto scheme_color_node = gs_node.append_child("a:schemeClr");
    scheme_color_node.append_attribute("val").set_value("phClr");
    scheme_color_node.append_child("a:tint").append_attribute("val").set_value(50000);
    scheme_color_node.append_child("a:satMod").append_attribute("val").set_value(300000);
    
    gs_node = grad_fill_list.append_child("a:gs");
    gs_node.append_attribute("pos").set_value(35000);
    scheme_color_node = gs_node.append_child("a:schemeClr");
    scheme_color_node.append_attribute("val").set_value("phClr");
    scheme_color_node.append_child("a:tint").append_attribute("val").set_value(37000);
    scheme_color_node.append_child("a:satMod").append_attribute("val").set_value(300000);
    
    gs_node = grad_fill_list.append_child("a:gs");
    gs_node.append_attribute("pos").set_value(100000);
    scheme_color_node = gs_node.append_child("a:schemeClr");
    scheme_color_node.append_attribute("val").set_value("phClr");
    scheme_color_node.append_child("a:tint").append_attribute("val").set_value(15000);
    scheme_color_node.append_child("a:satMod").append_attribute("val").set_value(350000);
    
    auto lin_node = grad_fill_node.append_child("a:lin");
    lin_node.append_attribute("ang").set_value(16200000);
    lin_node.append_attribute("scaled").set_value(1);
    
    grad_fill_node = fill_style_list_node.append_child("a:gradFill");
    grad_fill_node.append_attribute("rotWithShape").set_value(1);
    
    grad_fill_list = grad_fill_node.append_child("a:gsLst");
    gs_node = grad_fill_list.append_child("a:gs");
    gs_node.append_attribute("pos").set_value(0);
    scheme_color_node = gs_node.append_child("a:schemeClr");
    scheme_color_node.append_attribute("val").set_value("phClr");
    scheme_color_node.append_child("a:shade").append_attribute("val").set_value(51000);
    scheme_color_node.append_child("a:satMod").append_attribute("val").set_value(130000);
    
    gs_node = grad_fill_list.append_child("a:gs");
    gs_node.append_attribute("pos").set_value(80000);
    scheme_color_node = gs_node.append_child("a:schemeClr");
    scheme_color_node.append_attribute("val").set_value("phClr");
    scheme_color_node.append_child("a:shade").append_attribute("val").set_value(93000);
    scheme_color_node.append_child("a:satMod").append_attribute("val").set_value(130000);
    
    gs_node = grad_fill_list.append_child("a:gs");
    gs_node.append_attribute("pos").set_value(100000);
    scheme_color_node = gs_node.append_child("a:schemeClr");
    scheme_color_node.append_attribute("val").set_value("phClr");
    scheme_color_node.append_child("a:shade").append_attribute("val").set_value(94000);
    scheme_color_node.append_child("a:satMod").append_attribute("val").set_value(135000);
    
    lin_node = grad_fill_node.append_child("a:lin");
    lin_node.append_attribute("ang").set_value(16200000);
    lin_node.append_attribute("scaled").set_value(0);
    
    auto line_style_list_node = format_scheme_node.append_child("a:lnStyleLst");
    
    auto ln_node = line_style_list_node.append_child("a:ln");
    ln_node.append_attribute("w").set_value(9525);
    ln_node.append_attribute("cap").set_value("flat");
    ln_node.append_attribute("cmpd").set_value("sng");
    ln_node.append_attribute("algn").set_value("ctr");
    
    auto solid_fill_node = ln_node.append_child("a:solidFill");
    scheme_color_node = solid_fill_node.append_child("a:schemeClr");
    scheme_color_node.append_attribute("val").set_value("phClr");
    scheme_color_node.append_child("a:shade").append_attribute("val").set_value(95000);
    scheme_color_node.append_child("a:satMod").append_attribute("val").set_value(105000);
    ln_node.append_child("a:prstDash").append_attribute("val").set_value("solid");
    
    ln_node = line_style_list_node.append_child("a:ln");
    ln_node.append_attribute("w").set_value(25400);
    ln_node.append_attribute("cap").set_value("flat");
    ln_node.append_attribute("cmpd").set_value("sng");
    ln_node.append_attribute("algn").set_value("ctr");
    
    solid_fill_node = ln_node.append_child("a:solidFill");
    scheme_color_node = solid_fill_node.append_child("a:schemeClr");
    scheme_color_node.append_attribute("val").set_value("phClr");
    ln_node.append_child("a:prstDash").append_attribute("val").set_value("solid");
    
    ln_node = line_style_list_node.append_child("a:ln");
    ln_node.append_attribute("w").set_value(38100);
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
    outer_shadow_node.append_attribute("blurRad").set_value(40000);
    outer_shadow_node.append_attribute("dist").set_value(20000);
    outer_shadow_node.append_attribute("dir").set_value(5400000);
    outer_shadow_node.append_attribute("rotWithShape").set_value(0);
    auto srgb_clr_node = outer_shadow_node.append_child("a:srgbClr");
    srgb_clr_node.append_attribute("val").set_value("000000");
    srgb_clr_node.append_child("a:alpha").append_attribute("val").set_value(38000);
    
    effect_style_node = effect_style_list_node.append_child("a:effectStyle");
    effect_list_node = effect_style_node.append_child("a:effectLst");
    outer_shadow_node = effect_list_node.append_child("a:outerShdw");
    outer_shadow_node.append_attribute("blurRad").set_value(40000);
    outer_shadow_node.append_attribute("dist").set_value(23000);
    outer_shadow_node.append_attribute("dir").set_value(5400000);
    outer_shadow_node.append_attribute("rotWithShape").set_value(0);
    srgb_clr_node = outer_shadow_node.append_child("a:srgbClr");
    srgb_clr_node.append_attribute("val").set_value("000000");
    srgb_clr_node.append_child("a:alpha").append_attribute("val").set_value(35000);
    
    effect_style_node = effect_style_list_node.append_child("a:effectStyle");
    effect_list_node = effect_style_node.append_child("a:effectLst");
    outer_shadow_node = effect_list_node.append_child("a:outerShdw");
    outer_shadow_node.append_attribute("blurRad").set_value(40000);
    outer_shadow_node.append_attribute("dist").set_value(23000);
    outer_shadow_node.append_attribute("dir").set_value(5400000);
    outer_shadow_node.append_attribute("rotWithShape").set_value(0);
    srgb_clr_node = outer_shadow_node.append_child("a:srgbClr");
    srgb_clr_node.append_attribute("val").set_value("000000");
    srgb_clr_node.append_child("a:alpha").append_attribute("val").set_value(35000);
    auto scene3d_node = effect_style_node.append_child("a:scene3d");
    auto camera_node = scene3d_node.append_child("a:camera");
    camera_node.append_attribute("prst").set_value("orthographicFront");
    auto rot_node = camera_node.append_child("a:rot");
    rot_node.append_attribute("lat").set_value(0);
    rot_node.append_attribute("lon").set_value(0);
    rot_node.append_attribute("rev").set_value(0);
    auto light_rig_node = scene3d_node.append_child("a:lightRig");
    light_rig_node.append_attribute("rig").set_value("threePt");
    light_rig_node.append_attribute("dir").set_value("t");
    rot_node = light_rig_node.append_child("a:rot");
    rot_node.append_attribute("lat").set_value(0);
    rot_node.append_attribute("lon").set_value(0);
    rot_node.append_attribute("rev").set_value(1200000);
    
    auto bevel_node = effect_style_node.append_child("a:sp3d").append_child("a:bevelT");
    bevel_node.append_attribute("w").set_value(63500);
    bevel_node.append_attribute("h").set_value(25400);
    
    auto bg_fill_style_list_node = format_scheme_node.append_child("a:bgFillStyleLst");
    
    bg_fill_style_list_node.append_child("a:solidFill").append_child("a:schemeClr").append_attribute("val").set_value("phClr");
    
    grad_fill_node = bg_fill_style_list_node.append_child("a:gradFill");
    grad_fill_node.append_attribute("rotWithShape").set_value(1);
    
    grad_fill_list = grad_fill_node.append_child("a:gsLst");
    gs_node = grad_fill_list.append_child("a:gs");
    gs_node.append_attribute("pos").set_value(0);
    scheme_color_node = gs_node.append_child("a:schemeClr");
    scheme_color_node.append_attribute("val").set_value("phClr");
    scheme_color_node.append_child("a:tint").append_attribute("val").set_value(40000);
    scheme_color_node.append_child("a:satMod").append_attribute("val").set_value(350000);
    
    gs_node = grad_fill_list.append_child("a:gs");
    gs_node.append_attribute("pos").set_value(40000);
    scheme_color_node = gs_node.append_child("a:schemeClr");
    scheme_color_node.append_attribute("val").set_value("phClr");
    scheme_color_node.append_child("a:tint").append_attribute("val").set_value(45000);
    scheme_color_node.append_child("a:shade").append_attribute("val").set_value(99000);
    scheme_color_node.append_child("a:satMod").append_attribute("val").set_value(350000);
    
    gs_node = grad_fill_list.append_child("a:gs");
    gs_node.append_attribute("pos").set_value(100000);
    scheme_color_node = gs_node.append_child("a:schemeClr");
    scheme_color_node.append_attribute("val").set_value("phClr");
    scheme_color_node.append_child("a:shade").append_attribute("val").set_value(20000);
    scheme_color_node.append_child("a:satMod").append_attribute("val").set_value(255000);
    
    auto path_node = grad_fill_node.append_child("a:path");
    path_node.append_attribute("path").set_value("circle");
    auto fill_to_rect_node = path_node.append_child("a:fillToRect");
    fill_to_rect_node.append_attribute("l").set_value(50000);
    fill_to_rect_node.append_attribute("t").set_value(-80000);
    fill_to_rect_node.append_attribute("r").set_value(50000);
    fill_to_rect_node.append_attribute("b").set_value(180000);
    
    grad_fill_node = bg_fill_style_list_node.append_child("a:gradFill");
    grad_fill_node.append_attribute("rotWithShape").set_value(1);
    
    grad_fill_list = grad_fill_node.append_child("a:gsLst");
    gs_node = grad_fill_list.append_child("a:gs");
    gs_node.append_attribute("pos").set_value(0);
    scheme_color_node = gs_node.append_child("a:schemeClr");
    scheme_color_node.append_attribute("val").set_value("phClr");
    scheme_color_node.append_child("a:tint").append_attribute("val").set_value(80000);
    scheme_color_node.append_child("a:satMod").append_attribute("val").set_value(300000);
    
    gs_node = grad_fill_list.append_child("a:gs");
    gs_node.append_attribute("pos").set_value(100000);
    scheme_color_node = gs_node.append_child("a:schemeClr");
    scheme_color_node.append_attribute("val").set_value("phClr");
    scheme_color_node.append_child("a:shade").append_attribute("val").set_value(30000);
    scheme_color_node.append_child("a:satMod").append_attribute("val").set_value(200000);
    
    path_node = grad_fill_node.append_child("a:path");
    path_node.append_attribute("path").set_value("circle");
    fill_to_rect_node = path_node.append_child("a:fillToRect");
    fill_to_rect_node.append_attribute("l").set_value(50000);
    fill_to_rect_node.append_attribute("t").set_value(50000);
    fill_to_rect_node.append_attribute("r").set_value(50000);
    fill_to_rect_node.append_attribute("b").set_value(50000);
    
    theme_node.append_child("a:objectDefaults");
    theme_node.append_child("a:extraClrSchemeLst");
    
    std::stringstream ss;
    doc.print(ss);
    return ss.str();
}

std::string write_properties_app(const workbook &wb)
{
    pugi::xml_document doc;
    auto root_node = doc.append_child("Properties");
    root_node.append_attribute("xmlns").set_value("http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");
    root_node.append_attribute("xmlns:vt").set_value("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
    
    root_node.append_child("Application").text().set("Microsoft Excel");
    root_node.append_child("DocSecurity").text().set("0");
    root_node.append_child("ScaleCrop").text().set("false");
    root_node.append_child("Company");
    root_node.append_child("LinksUpToDate").text().set("false");
    root_node.append_child("SharedDoc").text().set("false");
    root_node.append_child("HyperlinksChanged").text().set("false");
    root_node.append_child("AppVersion").text().set("12.0000");
    
    auto heading_pairs_node = root_node.append_child("HeadingPairs");
    auto heading_pairs_vector_node = heading_pairs_node.append_child("vt:vector");
    heading_pairs_vector_node.append_attribute("baseType").set_value("variant");
    heading_pairs_vector_node.append_attribute("size").set_value("2");
    heading_pairs_vector_node.append_child("vt:variant").append_child("vt:lpstr").text().set("Worksheets");
    heading_pairs_vector_node.append_child("vt:variant").append_child("vt:i4").text().set(std::to_string(wb.get_sheet_names().size()).c_str());
    
    auto titles_of_parts_node = root_node.append_child("TitlesOfParts");
    auto titles_of_parts_vector_node = titles_of_parts_node.append_child("vt:vector");
    titles_of_parts_vector_node.append_attribute("baseType").set_value("lpstr");
    titles_of_parts_vector_node.append_attribute("size").set_value(std::to_string(wb.get_sheet_names().size()).c_str());
    
    for(auto ws : wb)
    {
        titles_of_parts_vector_node.append_child("vt:lpstr").text().set(ws.get_title().c_str());
    }
    
    std::stringstream ss;
    doc.save(ss);
    
    return ss.str();
}

std::string write_root_rels(const workbook &)
{
    std::vector<relationship> relationships;
    
    relationships.push_back(relationship(relationship::type::extended_properties, "rId3", "docProps/app.xml"));
    relationships.push_back(relationship(relationship::type::core_properties, "rId2", "docProps/core.xml"));
    relationships.push_back(relationship(relationship::type::office_document, "rId1", "xl/workbook.xml"));
    
    return write_relationships(relationships, "");
}

std::string write_workbook(const workbook &wb)
{
    std::size_t num_visible = 0;
    
    for(auto ws : wb)
    {
        if(ws.get_page_setup().get_sheet_state() == xlnt::page_setup::sheet_state::visible)
        {
            num_visible++;
        }
    }
    
    if(num_visible == 0)
    {
        throw value_error();
    }
    
    pugi::xml_document doc;
    auto root_node = doc.append_child("workbook");
    root_node.append_attribute("xmlns").set_value("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    root_node.append_attribute("xmlns:r").set_value("http://schemas.openxmlformats.org/officeDocument/2006/relationships");
    
    auto file_version_node = root_node.append_child("fileVersion");
    file_version_node.append_attribute("appName").set_value("xl");
    file_version_node.append_attribute("lastEdited").set_value("4");
    file_version_node.append_attribute("lowestEdited").set_value("4");
    file_version_node.append_attribute("rupBuild").set_value("4505");
    
    auto workbook_pr_node = root_node.append_child("workbookPr");
    workbook_pr_node.append_attribute("codeName").set_value("ThisWorkbook");
    workbook_pr_node.append_attribute("defaultThemeVersion").set_value("124226");
    workbook_pr_node.append_attribute("date1904").set_value(wb.get_properties().excel_base_date == calendar::mac_1904 ? 1 : 0);
    
    auto book_views_node = root_node.append_child("bookViews");
    auto workbook_view_node = book_views_node.append_child("workbookView");
    workbook_view_node.append_attribute("activeTab").set_value("0");
    workbook_view_node.append_attribute("autoFilterDateGrouping").set_value("1");
    workbook_view_node.append_attribute("firstSheet").set_value("0");
    workbook_view_node.append_attribute("minimized").set_value("0");
    workbook_view_node.append_attribute("showHorizontalScroll").set_value("1");
    workbook_view_node.append_attribute("showSheetTabs").set_value("1");
    workbook_view_node.append_attribute("showVerticalScroll").set_value("1");
    workbook_view_node.append_attribute("tabRatio").set_value("600");
    workbook_view_node.append_attribute("visibility").set_value("visible");
    
    auto sheets_node = root_node.append_child("sheets");
    auto defined_names_node = root_node.append_child("definedNames");
    
    for(auto relationship : wb.get_relationships())
    {
        if(relationship.get_type() == relationship::type::worksheet)
        {
            std::string sheet_index_string = relationship.get_target_uri();
            sheet_index_string = sheet_index_string.substr(0, sheet_index_string.find('.'));
            sheet_index_string = sheet_index_string.substr(sheet_index_string.find_last_of('/'));
            auto iter = sheet_index_string.end();
            iter--;
            while (isdigit(*iter)) iter--;
            auto first_digit = iter - sheet_index_string.begin();
            sheet_index_string = sheet_index_string.substr(static_cast<std::string::size_type>(first_digit + 1));
            std::size_t sheet_index = static_cast<std::size_t>(std::stoll(sheet_index_string) - 1);
            
            auto ws = wb.get_sheet_by_index(sheet_index);
            
            auto sheet_node = sheets_node.append_child("sheet");
            sheet_node.append_attribute("name").set_value(ws.get_title().c_str());
            sheet_node.append_attribute("r:id").set_value(relationship.get_id().c_str());
            sheet_node.append_attribute("sheetId").set_value(std::to_string(sheet_index + 1).c_str());
            
            if(ws.has_auto_filter())
            {
                auto defined_name_node = defined_names_node.append_child("definedName");
                defined_name_node.append_attribute("name").set_value("_xlnm._FilterDatabase");
                defined_name_node.append_attribute("hidden").set_value(1);
                defined_name_node.append_attribute("localSheetId").set_value(0);
                std::string name = "'" + ws.get_title() + "'!" + range_reference::make_absolute(ws.get_auto_filter()).to_string();
                defined_name_node.text().set(name.c_str());
            }
        }
    }
    
    auto calc_pr_node = root_node.append_child("calcPr");
    calc_pr_node.append_attribute("calcId").set_value("124519");
    calc_pr_node.append_attribute("calcMode").set_value("auto");
    calc_pr_node.append_attribute("fullCalcOnLoad").set_value("1");
    
    std::stringstream ss;
    doc.save(ss);
    
    return ss.str();
}

std::string workbook_serializer::write_workbook_rels(const workbook &wb)
{
    return write_relationships(wb.get_relationships(), "xl/");
}
    
std::string workbook_serializer::write_defined_names(const xlnt::workbook &wb)
{
    pugi::xml_document doc;
    auto names = doc.root().append_child("names");
    
    for(auto named_range : wb.get_named_ranges())
    {
        names.append_child(named_range.get_name().c_str());
    }
    
    std::ostringstream stream;
    doc.save(stream);
    
    return stream.str();
}

} // namespace xlnt
