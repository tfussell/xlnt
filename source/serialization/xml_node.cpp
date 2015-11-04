#include <xlnt/serialization/xml_node.hpp>
#include <xlnt/serialization/xml_serializer.hpp>

#include <detail/xml_node_impl.hpp>

namespace xlnt {

xml_node::xml_node() : d_(new detail::xml_node_impl)
{
}

xml_node::xml_node(const detail::xml_node_impl &d) : xml_node()
{
    d_->node = d.node;
}

xml_node::~xml_node()
{
}

xml_node::xml_node(const xml_node &other) : xml_node()
{
    d_->node = other.d_->node;
}

xml_node &xml_node::operator=(const xlnt::xml_node &other)
{
    d_->node = other.d_->node;

    return *this;
}

string xml_node::get_name() const
{
    return d_->node.name();
}

void xml_node::set_name(const string &name)
{
    d_->node.set_name(name.data());
}

bool xml_node::has_text() const
{
    return d_->node.text() != nullptr;
}

string xml_node::get_text() const
{
    return d_->node.text().as_string();
}

void xml_node::set_text(const string &text)
{
    d_->node.text().set(text.data());
}

const std::vector<xml_node> xml_node::get_children() const
{
    std::vector<xml_node> children;

    for (auto child : d_->node.children())
    {
        children.push_back(xml_node(detail::xml_node_impl(child)));
    }

    return children;
}

xml_node xml_node::add_child(const xml_node &child)
{
    auto child_node = xml_node(detail::xml_node_impl(d_->node.append_child(child.get_name().data())));

    for (auto attr : child.get_attributes())
    {
        child_node.add_attribute(attr.first, attr.second);
    }

    for (auto child_child : child.get_children())
    {
        child_node.add_child(child_child);
    }

    return child_node;
}

xml_node xml_node::add_child(const string &child_name)
{
    return xml_node(detail::xml_node_impl(d_->node.append_child(child_name.data())));
}

const std::vector<xml_node::string_pair> xml_node::get_attributes() const
{
    std::vector<string_pair> attributes;

    for (auto attr : d_->node.attributes())
    {
        attributes.push_back(std::make_pair<string, string>(attr.name(), attr.value()));
    }

    return attributes;
}

void xml_node::add_attribute(const string &name, const string &value)
{
    d_->node.append_attribute(name.data()).set_value(value.data());
}

bool xml_node::has_attribute(const string &attribute_name) const
{
    return d_->node.attribute(attribute_name.data()) != nullptr;
}

string xml_node::get_attribute(const string &attribute_name) const
{
    return d_->node.attribute(attribute_name.data()).value();
}

bool xml_node::has_child(const string &child_name) const
{
    return d_->node.child(child_name.data()) != nullptr;
}

xml_node xml_node::get_child(const string &child_name)
{
    return xml_node(detail::xml_node_impl(d_->node.child(child_name.data())));
}

const xml_node xml_node::get_child(const string &child_name) const
{
    return xml_node(detail::xml_node_impl(d_->node.child(child_name.data())));
}

string xml_node::to_string() const
{
    return xml_serializer::serialize_node(*this);
}

} // namespace xlnt
