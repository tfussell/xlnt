#pragma once

#include <memory>
#include <string>
#include <vector>

namespace xlnt {
namespace detail {
struct xml_node_impl;
} // namespace detail
    
class xml_document;
    
class xml_node
{
public:
    using string_pair = std::pair<std::string, std::string>;
    
    xml_node();
    xml_node(const xml_node &other);
    ~xml_node();
    
    xml_node &operator=(const xml_node &other);
    
    std::string get_name() const;
    void set_name(const std::string &name);
    
    bool has_text() const;
    std::string get_text() const;
    void set_text(const std::string &text);
    
    const std::vector<xml_node> get_children() const;
    bool has_child(const std::string &child_name) const;
    xml_node get_child(const std::string &child_name);
    const xml_node get_child(const std::string &child_name) const;
    xml_node add_child(const xml_node &child);
    xml_node add_child(const std::string &child_name);
    
    const std::vector<string_pair> get_attributes() const;
    bool has_attribute(const std::string &attribute_name) const;
    std::string get_attribute(const std::string &attribute_name) const;
    void add_attribute(const std::string &name, const std::string &value);
    
    std::string to_string() const;
    
private:
    friend class xml_document;
    friend class xml_serializer;
    xml_node(const detail::xml_node_impl &d);
    std::unique_ptr<detail::xml_node_impl> d_;
};

} // namespace xlnt
