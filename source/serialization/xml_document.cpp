#include <xlnt/serialization/xml_document.hpp>
#include <xlnt/serialization/xml_node.hpp>
#include <xlnt/serialization/xml_serializer.hpp>

#include <detail/include_pugixml.hpp>
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

xml_document::xml_document(xml_document &&other)
{
	std::swap(d_, other.d_);
}

xml_document::~xml_document()
{
}

void xml_document::set_encoding(const string &encoding)
{
    d_->encoding = encoding;
}

void xml_document::add_namespace(const string &id, const string &uri)
{
    d_->doc.first_child().append_attribute((id.empty() ? "xmlns" : "xmlns:" + id).data()).set_value(uri.data());
}

xml_node xml_document::add_child(const xml_node &child)
{
    auto child_node = d_->doc.root().append_copy(child.d_->node);
    return xml_node(detail::xml_node_impl(child_node));
}

xml_node xml_document::add_child(const string &child_name)
{
    auto child = d_->doc.root().append_child(child_name.data());
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

string xml_document::to_string() const
{
    return xml_serializer::serialize(*this);
}

xml_document &xml_document::from_string(const string &xml_string)
{
	d_->doc.load(xml_string.data());

    return *this;
}

xml_node xml_document::get_child(const string &child_name)
{
    return xml_node(detail::xml_node_impl(d_->doc.child(child_name.data())));
}

const xml_node xml_document::get_child(const string &child_name) const
{
    return xml_node(detail::xml_node_impl(d_->doc.child(child_name.data())));
}

} // namespace xlnt
