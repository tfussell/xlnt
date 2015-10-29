#include <xlnt/s11n/excel_serializer.hpp>
#include <xlnt/s11n/manifest_serializer.hpp>
#include <xlnt/workbook/manifest.hpp>
#include <xlnt/workbook/workbook.hpp>

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

xlnt::workbook load_workbook(xlnt::zip_file &archive, bool guess_types, bool data_only)
{
    xlnt::workbook wb;
    
    wb.set_guess_types(guess_types);
    wb.set_data_only(data_only);
    
    auto content_types = xlnt::read_content_types(archive);
    auto type = xlnt::determine_document_type(content_types);
    
    if(type != "excel")
    {
        throw xlnt::invalid_file_exception("");
    }
    
    wb.clear();
    
    auto workbook_relationships = read_relationships(archive, "xl/workbook.xml");
    
    for(auto relationship : workbook_relationships)
    {
        wb.create_relationship(relationship.get_id(), relationship.get_target_uri(), relationship.get_type());
    }
    
    pugi::xml_document doc;
    doc.load(archive.read("xl/workbook.xml").c_str());
    
    auto root_node = doc.child("workbook");
    
    auto workbook_pr_node = root_node.child("workbookPr");
    wb.get_properties().excel_base_date = (workbook_pr_node.attribute("date1904") != nullptr && workbook_pr_node.attribute("date1904").as_int() != 0) ? xlnt::calendar::mac_1904 : xlnt::calendar::windows_1900;
    
    auto sheets_node = root_node.child("sheets");
    
    xlnt::shared_strings_reader shared_strings_reader_;
    auto shared_strings = shared_strings_reader_.read_strings(archive);
    
    xlnt::style_reader style_reader_(wb);
    style_reader_.read_styles(archive);
    
    for(const auto &border_ : style_reader_.get_borders())
    {
        wb.add_border(border_);
    }
    
    for(const auto &fill_ : style_reader_.get_fills())
    {
        wb.add_fill(fill_);
    }
    
    for(const auto &font_ : style_reader_.get_fonts())
    {
        wb.add_font(font_);
    }
    
    for(const auto &number_format_ : style_reader_.get_number_formats())
    {
        wb.add_number_format(number_format_);
    }
    
    for(auto &color_rgb : style_reader_.get_colors())
    {
        wb.add_color(xlnt::color(xlnt::color::type::rgb, color_rgb));
    }
    
    for(const auto &style : style_reader_.get_styles())
    {
        wb.add_style(style);
    }
    
    for(auto sheet_node : sheets_node.children("sheet"))
    {
        auto rel = wb.get_relationship(sheet_node.attribute("r:id").as_string());
        auto ws = wb.create_sheet(sheet_node.attribute("name").as_string(), rel);
        
        xlnt::read_worksheet(ws, archive, rel, shared_strings);
    }
    
    return wb;
}
    
} // namespace

namespace xlnt {

std::string excel_reader::CentralDirectorySignature()
{
    return "\x50\x4b\x05\x06";
}

std::string excel_reader::repair_central_directory(const std::string &original)
{
    auto pos = find_string_in_string(original, CentralDirectorySignature());
    
    if(pos != std::string::npos)
    {
        return original.substr(0, pos + 22);
    }
    
    return original;
}

workbook excel_reader::load_workbook(std::istream &stream, bool guess_types, bool data_only)
{
    std::vector<std::uint8_t> bytes((std::istream_iterator<char>(stream)),
                                    std::istream_iterator<char>());
    
    return load_workbook(bytes, guess_types, data_only);
}

workbook excel_reader::load_workbook(const std::string &filename, bool guess_types, bool data_only)
{
    xlnt::zip_file archive;
    
    try
    {
        archive.load(filename);
    }
    catch(std::runtime_error)
    {
        throw invalid_file_exception(filename);
    }
    
    return ::load_workbook(archive, guess_types, data_only);
}

xlnt::workbook excel_reader::load_workbook(const std::vector<std::uint8_t> &bytes, bool guess_types, bool data_only)
{
    xlnt::zip_file archive;
    archive.load(bytes);
    
    return ::load_workbook(archive, guess_types, data_only);
}
    
excel_writer::excel_writer(workbook &wb) : wb_(wb)
{
}

void excel_writer::save(const std::string &filename, bool as_template)
{
    zip_file archive;
    write_data(archive, as_template);
    archive.save(filename);
}

void excel_writer::write_data(zip_file &archive, bool as_template)
{
    archive.writestr(constants::ArcRootRels, write_root_rels(wb_));
    archive.writestr(constants::ArcWorkbookRels, write_workbook_rels(wb_));
    archive.writestr(constants::ArcApp, write_properties_app(wb_));
    archive.writestr(constants::ArcCore, write_properties_core(wb_.get_properties()));
    
    if(wb_.has_loaded_theme())
    {
        archive.writestr(constants::ArcTheme, wb_.get_loaded_theme());
    }
    else
    {
        archive.writestr(constants::ArcTheme, write_theme());
    }
    
    archive.writestr(constants::ArcWorkbook, write_workbook(wb_));
    
    auto shared_strings = extract_all_strings(wb_);
    
    write_charts(archive);
    write_images(archive);
    write_shared_strings(archive, shared_strings);
    write_worksheets(archive, shared_strings);
    write_chartsheets(archive);
    write_external_links(archive);
    
    style_writer style_writer_(wb_);
    archive.writestr(constants::ArcStyles, style_writer_.write_table());
    
    auto manifest = write_content_types(wb_, as_template);
    archive.writestr(constants::ArcContentTypes, manifest);
}

void excel_writer::write_shared_strings(xlnt::zip_file &archive, const std::vector<std::string> &shared_strings)
{
    archive.writestr(constants::ArcSharedString, ::xlnt::write_shared_strings(shared_strings));
}

void excel_writer::write_images(zip_file &/*archive*/)
{
    
}

void excel_writer::write_charts(zip_file &/*archive*/)
{
    
}

void excel_writer::write_chartsheets(zip_file &/*archive*/)
{
    
}

void excel_writer::write_worksheets(zip_file &archive, const std::vector<std::string> &shared_strings)
{
    std::size_t index = 0;
    
    for(auto ws : wb_)
    {
        for(auto relationship : wb_.get_relationships())
        {
            if(relationship.get_type() == relationship::type::worksheet &&
               workbook::index_from_ws_filename(relationship.get_target_uri()) == index)
            {
                archive.writestr(relationship.get_target_uri(), write_worksheet(ws, shared_strings));
                break;
            }
        }
        
        index++;
    }
}

void excel_writer::write_external_links(zip_file &/*archive*/)
{
    
}

bool save_workbook(workbook &wb, const std::string &filename, bool as_template)
{
    excel_writer writer(wb);
    writer.save(filename, as_template);
    return true;
}

std::vector<std::uint8_t> save_virtual_workbook(xlnt::workbook &wb, bool as_template)
{
    zip_file archive;
    excel_writer writer(wb);
    writer.write_data(archive, as_template);
    std::vector<std::uint8_t> buffer;
    archive.save(buffer);
    
    return buffer;
}
    
} // namespace xlnt
