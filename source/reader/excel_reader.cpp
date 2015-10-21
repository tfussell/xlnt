#include <xlnt/reader/excel_reader.hpp>
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

    
} // namespace

namespace xlnt {

std::string CentralDirectorySignature()
{
    return "\x50\x4b\x05\x06";
}

std::string repair_central_directory(const std::string &original)
{
    auto pos = find_string_in_string(original, CentralDirectorySignature());
    
    if(pos != std::string::npos)
    {
        return original.substr(0, pos + 22);
    }
    
    return original;
}

workbook load_workbook(const std::string &filename, bool guess_types, bool data_only)
{
    workbook wb;
    
    wb.set_guess_types(guess_types);
    wb.set_data_only(data_only);
    wb.load(filename);
    
    return wb;
}

xlnt::workbook load_workbook(const std::vector<std::uint8_t> &bytes, bool guess_types, bool data_only)
{
    xlnt::workbook wb;

    wb.set_guess_types(guess_types);
    wb.set_data_only(data_only);
    wb.load(bytes);
    
    return wb;
}
    
} // namespace xlnt
