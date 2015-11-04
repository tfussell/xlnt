#pragma once

#include <memory>
#include <vector>

#include <xlnt/utils/string.hpp>

#include "xlnt_config.hpp"

namespace xlnt {
namespace detail { struct xml_document_impl; }

class xml_node;
class xml_serializer;

class XLNT_CLASS xml_document
{
  public:
    using string_pair = std::pair<string, string>;

    xml_document();
    xml_document(const xml_document &other);
    xml_document(xml_document &&other);
    ~xml_document();

    xml_document &operator=(const xml_document &other);
    xml_document &operator=(xml_document &&other);

    void set_encoding(const string &encoding);
    void add_namespace(const string &id, const string &uri);

    xml_node add_child(const xml_node &child);
    xml_node add_child(const string &child_name);

    xml_node get_root();
    const xml_node get_root() const;

    xml_node get_child(const string &child_name);
    const xml_node get_child(const string &child_name) const;

    string to_string() const;
    xml_document &from_string(const string &xml_string);

  private:
    friend class xml_serializer;
    std::unique_ptr<detail::xml_document_impl> d_;
};

} // namespace xlnt
