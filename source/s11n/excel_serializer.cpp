#include <xlnt/s11n/excel_serializer.hpp>
#include <xlnt/common/exceptions.hpp>
#include <xlnt/s11n/manifest_serializer.hpp>
#include <xlnt/s11n/relationship_serializer.hpp>
#include <xlnt/s11n/shared_strings_serializer.hpp>
#include <xlnt/s11n/style_serializer.hpp>
#include <xlnt/s11n/theme_serializer.hpp>
#include <xlnt/s11n/workbook_serializer.hpp>
#include <xlnt/s11n/worksheet_serializer.hpp>
#include <xlnt/s11n/xml_document.hpp>
#include <xlnt/s11n/xml_serializer.hpp>
#include <xlnt/workbook/document_properties.hpp>
#include <xlnt/workbook/manifest.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/worksheet/worksheet.hpp>

#include <detail/constants.hpp>

namespace {
    
std::string::size_type find_string_in_string(const std::string &string, const std::string &substring)
{
    std::string::size_type possible_match_index = string.find(substring.at(0));
    
    while(possible_match_index != std::string::npos)
    {
        if(string.substr(possible_match_index, substring.size()) == substring)
        {
            return possible_match_index;
        }
        
        possible_match_index = string.find(substring.at(0), possible_match_index + 1);
    }
    
    return possible_match_index;
}

bool load_workbook(xlnt::zip_file &archive, bool guess_types, bool data_only, xlnt::workbook &wb)
{
    wb.set_guess_types(guess_types);
    wb.set_data_only(data_only);
    
    xlnt::manifest_serializer ms(wb.get_manifest());
    ms.read_manifest(xlnt::xml_serializer::deserialize(archive.read(xlnt::constants::ArcContentTypes)));
    
    if(xlnt::workbook_serializer::determine_document_type(wb.get_manifest()) != "excel")
    {
        throw xlnt::invalid_file_exception("");
    }
    
    wb.clear();
    
    std::vector<xlnt::relationship> workbook_relationships;
    xlnt::relationship_serializer::read_relationships(xlnt::xml_serializer::deserialize(archive.read("xl/_rels/workbook.xml.rels")), "", workbook_relationships);
    
    for(auto relationship : workbook_relationships)
    {
        wb.create_relationship(relationship.get_id(), relationship.get_target_uri(), relationship.get_type());
    }
    
    auto xml = xlnt::xml_serializer::deserialize(archive.read(xlnt::constants::ArcWorkbook));
    
    auto &root_node = xml.root();
    
    auto &workbook_pr_node = root_node.get_child("workbookPr");
    wb.get_properties().excel_base_date = (workbook_pr_node.has_attribute("date1904") && workbook_pr_node.get_attribute("date1904") != "0") ? xlnt::calendar::mac_1904 : xlnt::calendar::windows_1900;
    
    xlnt::shared_strings_serializer shared_strings_serializer_;
    std::vector<std::string> shared_strings;
    shared_strings_serializer_.read_strings(xlnt::xml_serializer::deserialize(archive.read(xlnt::constants::ArcSharedString)), shared_strings);
    
    xlnt::style_serializer style_reader_(wb);
    style_reader_.read_stylesheet(xlnt::xml_serializer::deserialize(archive.read(xlnt::constants::ArcStyles)));
    
    auto &sheets_node = root_node.get_child("sheets");
    
    for(auto sheet_node : sheets_node.get_children())
    {
        auto rel = wb.get_relationship(sheet_node.get_attribute("r:id"));
        auto ws = wb.create_sheet(sheet_node.get_attribute("name"), rel);
        
        xlnt::worksheet_serializer worksheet_serializer(ws);
        worksheet_serializer.read_worksheet(xlnt::xml_serializer::deserialize(archive.read(rel.get_target_uri())), shared_strings, rel);
    }
    
    return true;
}
    
} // namespace

namespace xlnt {

const std::string excel_serializer::central_directory_signature()
{
    return "\x50\x4b\x05\x06";
}

std::string excel_serializer::repair_central_directory(const std::string &original)
{
    auto pos = find_string_in_string(original, central_directory_signature());
    
    if(pos != std::string::npos)
    {
        return original.substr(0, pos + 22);
    }
    
    return original;
}

bool excel_serializer::load_stream_workbook(std::istream &stream, bool guess_types, bool data_only)
{
    std::vector<std::uint8_t> bytes((std::istream_iterator<char>(stream)),
                                    std::istream_iterator<char>());

    return load_virtual_workbook(bytes, guess_types, data_only);
}

bool excel_serializer::load_workbook(const std::string &filename, bool guess_types, bool data_only)
{
    try
    {
        archive_.load(filename);
    }
    catch(std::runtime_error)
    {
        throw invalid_file_exception(filename);
    }
    
    return ::load_workbook(archive_, guess_types, data_only, wb_);
}

bool excel_serializer::load_virtual_workbook(const std::vector<std::uint8_t> &bytes, bool guess_types, bool data_only)
{
    archive_.load(bytes);
    
    return ::load_workbook(archive_, guess_types, data_only, wb_);
}
    
excel_serializer::excel_serializer(workbook &wb) : wb_(wb)
{
}

void excel_serializer::write_data(bool as_template)
{
    relationship_serializer relationship_serializer_;
    
    xlnt::xml_document root_rels_xml;
    relationship_serializer_.write_relationships(wb_.get_root_relationships(), "", root_rels_xml);
    archive_.writestr(constants::ArcRootRels, xml_serializer::serialize(root_rels_xml));
    
    xml_document workbook_rels_xml;
    relationship_serializer_.write_relationships(wb_.get_relationships(), "", workbook_rels_xml);
    archive_.writestr(constants::ArcWorkbookRels, xml_serializer::serialize(workbook_rels_xml));
    
    xml_document properties_app_xml;
    workbook_serializer workbook_serializer_(wb_);
    archive_.writestr(constants::ArcApp, xml_serializer::serialize(workbook_serializer_.write_properties_app()));
    archive_.writestr(constants::ArcCore, xml_serializer::serialize(workbook_serializer_.write_properties_core()));
    
    theme_serializer theme_serializer_;
    xml_document theme_xml = theme_serializer_.write_theme(wb_.get_loaded_theme());
    archive_.writestr(constants::ArcTheme, xml_serializer::serialize(theme_xml));
    
    archive_.writestr(constants::ArcWorkbook, xml_serializer::serialize(workbook_serializer_.write_workbook()));
    
    style_serializer style_serializer_(wb_);
    xml_document style_xml;
    style_serializer_.write_stylesheet(style_xml);
    archive_.writestr(constants::ArcStyles, xml_serializer::serialize(style_xml));
    
    manifest_serializer manifest_serializer_(wb_.get_manifest());
    xml_document manifest_xml;
    manifest_serializer_.write_manifest(manifest_xml);
    archive_.writestr(constants::ArcContentTypes, xml_serializer::serialize(manifest_xml));
    
    write_worksheets();
}

void excel_serializer::write_worksheets()
{
    std::size_t index = 0;
    
    for(auto ws : wb_)
    {
        for(auto relationship : wb_.get_relationships())
        {
            if(relationship.get_type() == relationship::type::worksheet &&
               workbook::index_from_ws_filename(relationship.get_target_uri()) == index)
            {
                worksheet_serializer serializer_(ws);
                xml_document xml;
                serializer_.write_worksheet(shared_strings_, xml);
                archive_.writestr(relationship.get_target_uri(),  xml_serializer::serialize(xml));
                break;
            }
        }
        
        index++;
    }
}

void excel_serializer::write_external_links()
{
    
}
    
bool excel_serializer::save_stream_workbook(std::ostream &stream, bool as_template)
{
    write_data(as_template);
    archive_.save(stream);
    
    return true;
}

bool excel_serializer::save_workbook(const std::string &filename, bool as_template)
{
    write_data(as_template);
    archive_.save(filename);

    return true;
}

bool excel_serializer::save_virtual_workbook(std::vector<std::uint8_t> &bytes, bool as_template)
{
    write_data(as_template);
    archive_.save(bytes);
    
    return true;
}
    
} // namespace xlnt
