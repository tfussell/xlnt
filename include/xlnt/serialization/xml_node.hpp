#pragma once

#include <memory>
#include <vector>

#include <xlnt/utils/string.hpp>

#include "xlnt_config.hpp"

namespace xlnt {
namespace detail { struct xml_node_impl; }

class xml_document;

class XLNT_CLASS xml_node
{
  public:
    using string_pair = std::pair<string, string>;

    xml_node();
    xml_node(const xml_node &other);
    ~xml_node();

    xml_node &operator=(const xml_node &other);

    string get_name() const;
    void set_name(const string &name);

    bool has_text() const;
    string get_text() const;
    void set_text(const string &text);

    const std::vector<xml_node> get_children() const;
    bool has_child(const string &child_name) const;
    xml_node get_child(const string &child_name);
    const xml_node get_child(const string &child_name) const;
    xml_node add_child(const xml_node &child);
    xml_node add_child(const string &child_name);

    const std::vector<string_pair> get_attributes() const;
    bool has_attribute(const string &attribute_name) const;
    string get_attribute(const string &attribute_name) const;
    void add_attribute(const string &name, const string &value);

    string to_string() const;

  private:
    friend class xml_document;
    friend class xml_serializer;
    xml_node(const detail::xml_node_impl &d);
    std::unique_ptr<detail::xml_node_impl> d_;
};

} // namespace xlnt
