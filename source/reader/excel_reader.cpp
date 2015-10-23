#include <xlnt/common/datetime.hpp>
#include <xlnt/common/exceptions.hpp>
#include <xlnt/common/zip_file.hpp>
#include <xlnt/styles/style.hpp>
#include <xlnt/reader/excel_reader.hpp>
#include <xlnt/reader/shared_strings_reader.hpp>
#include <xlnt/reader/style_reader.hpp>
#include <xlnt/reader/workbook_reader.hpp>
#include <xlnt/reader/worksheet_reader.hpp>
#include <xlnt/workbook/document_properties.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/worksheet/worksheet.hpp>

#include "detail/include_pugixml.hpp"

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
    
} // namespace xlnt
