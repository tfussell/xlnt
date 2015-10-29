#pragma once

#include <string>
#include <vector>

#include <xlnt/s11n/xml_node.hpp>

namespace xlnt {

class xml_document
{
public:
    void set_encoding(const std::string &encoding);
    void add_namespace(const std::string &id, const std::string &uri);
    
    xml_node &root();
    const xml_node &root() const;
    
private:
    xml_node root_;
    std::string encoding_;
    std::vector<string_pair> namespaces_;
};

} // namespace xlnt
