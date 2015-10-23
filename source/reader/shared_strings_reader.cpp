#include <xlnt/common/zip_file.hpp>
#include <xlnt/reader/shared_strings_reader.hpp>

#include "detail/include_pugixml.hpp"

namespace xlnt {

std::vector<std::string> shared_strings_reader::read_strings(zip_file &archive)
{
    static const std::string shared_strings_filename = "xl/sharedStrings.xml";
    
    if(!archive.has_file(shared_strings_filename))
    {
        return { };
    }
    
    pugi::xml_document doc;
    doc.load(archive.read(shared_strings_filename).c_str());
    
    auto root_node = doc.child("sst");
    int unique_count = root_node.attribute("uniqueCount").as_int();
 
    std::vector<std::string> shared_strings;
    
    for(auto si_node : root_node)
    {
        shared_strings.push_back(si_node.child("t").text().as_string());
    }
    
    if(unique_count != (int)shared_strings.size())
    {
        throw std::runtime_error("counts don't match");
    }
    
    return shared_strings;
}

} // namespace xlnt
