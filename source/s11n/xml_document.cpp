#include <xlnt/s11n/xml_document.hpp>
#include <xlnt/s11n/xml_node.hpp>
#include <xlnt/s11n/xml_serializer.hpp>

#include <detail/xml_document_impl.hpp>
#include <detail/xml_node_impl.hpp>

namespace xlnt {
    
xml_document::xml_document() : d_(new detail::xml_document_impl())
{
}

xml_document::xml_document(const xml_document &other) : xml_document()
{
    d_->doc.append_copy(other.d_->doc.root());
}

xml_document::~xml_document()
{
}

void xml_document::set_encoding(const std::string &encoding)
{
    d_->encoding = encoding;
}

void xml_document::add_namespace(const std::string &id, const std::string &uri)
{
    d_->doc.first_child().append_attribute((id.empty() ? "xmlns" : "xmlns:" + id).c_str()).set_value(uri.c_str());
}

xml_node xml_document::add_child(const xml_node &child)
{
    auto child_node = d_->doc.root().append_copy(child.d_->node);
    return xml_node(detail::xml_node_impl(child_node));
}

xml_node xml_document::add_child(const std::string &child_name)
{
    auto child = d_->doc.root().append_child(child_name.c_str());
    return xml_node(detail::xml_node_impl(child));
}

xml_node xml_document::get_root()
{
    return xml_node(detail::xml_node_impl(d_->doc.root()));
}

const xml_node xml_document::get_root() const
{
    return xml_node(detail::xml_node_impl(d_->doc.root()));
}

std::string xml_document::to_string() const
{
    return xml_serializer::serialize(*this);
}

xml_document &xml_document::from_string(const std::string &xml_string)
{
    auto doc = xml_serializer::deserialize(xml_string);
    std::swap(doc.d_, d_);
    
    return *this;
}

xml_node xml_document::get_child(const std::string &child_name)
{
    return xml_node(detail::xml_node_impl(d_->doc.child(child_name.c_str())));
}

const xml_node xml_document::get_child(const std::string &child_name) const
{
    return xml_node(detail::xml_node_impl(d_->doc.child(child_name.c_str())));
}

} // namespace xlnt