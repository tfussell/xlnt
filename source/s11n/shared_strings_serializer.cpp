#include <xlnt/s11n/shared_strings_serializer.hpp>
#include <xlnt/s11n/xml_document.hpp>
#include <xlnt/s11n/xml_node.hpp>

namespace xlnt {

bool shared_strings_serializer::read_strings(const xml_document &xml, std::vector<std::string> &strings)
{
    strings.clear();
    
    auto root_node = xml.root();
    root_node.set_name("sst");
    
    auto unique_count = std::stoull(root_node.get_attribute("uniqueCount"));
    
    for(const auto &si_node : root_node.get_children())
    {
        strings.push_back(si_node.get_child("t").get_text());
    }
    
    if(unique_count != strings.size())
    {
        throw std::runtime_error("counts don't match");
    }
    
    return true;
}

} // namespace xlnt
