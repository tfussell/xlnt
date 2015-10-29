#pragma once

#include <string>
#include <vector>

namespace xlnt {

using string_pair = std::pair<std::string, std::string>;

class xml_node
{
public:
    xml_node();
    explicit xml_node(const std::string &name);
    
    std::string get_name() const;
    void set_name(const std::string &name);
    
    bool has_text() const;
    std::string get_text() const;
    void set_text(const std::string &text);
    
    const std::vector<xml_node> &get_children() const;
    bool has_child(const std::string &child_name) const;
    const xml_node &get_child(const std::string &child_name) const;
    xml_node &add_child(const xml_node &child);
    xml_node &add_child(const std::string &child_name);
    
    const std::vector<string_pair> &get_attributes() const;
    bool has_attribute(const std::string &attribute_name) const;
    std::string get_attribute(const std::string &attribute_name) const;
    void add_attribute(const std::string &name, const std::string &value);
    
private:
    std::string name_;
    bool has_text_ = false;
    std::string text_;
    std::vector<string_pair> attributes_;
    std::vector<xml_node> children_;
};

} // namespace xlnt
