#pragma once

#include <memory>
#include <string>
#include <vector>

namespace xlnt {
namespace detail {
struct xml_document_impl;
} // namespace detail

class xml_node;
class xml_serializer;

class xml_document
{
  public:
    using string_pair = std::pair<std::string, std::string>;

    xml_document();
    xml_document(const xml_document &other);
    xml_document(xml_document &&other);
    ~xml_document();

    xml_document &operator=(const xml_document &other);
    xml_document &operator=(xml_document &&other);

    void set_encoding(const std::string &encoding);
    void add_namespace(const std::string &id, const std::string &uri);

    xml_node add_child(const xml_node &child);
    xml_node add_child(const std::string &child_name);

    xml_node get_root();
    const xml_node get_root() const;

    xml_node get_child(const std::string &child_name);
    const xml_node get_child(const std::string &child_name) const;

    std::string to_string() const;
    xml_document &from_string(const std::string &xml_string);

  private:
    friend class xml_serializer;
    std::unique_ptr<detail::xml_document_impl> d_;
};

} // namespace xlnt
