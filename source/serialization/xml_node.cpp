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

std::string xml_node::get_name() const
{
    return d_->node.name();
}

void xml_node::set_name(const std::string &name)
{
    d_->node.set_name(name.c_str());
}

bool xml_node::has_text() const
{
    return d_->node.text() != nullptr;
}

std::string xml_node::get_text() const
{
    return d_->node.text().as_string();
}

void xml_node::set_text(const std::string &text)
{
    d_->node.text().set(text.c_str());
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
    auto child_node = xml_node(detail::xml_node_impl(d_->node.append_child(child.get_name().c_str())));

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

xml_node xml_node::add_child(const std::string &child_name)
{
    return xml_node(detail::xml_node_impl(d_->node.append_child(child_name.c_str())));
}

const std::vector<xml_node::string_pair> xml_node::get_attributes() const
{
    std::vector<string_pair> attributes;

    for (auto attr : d_->node.attributes())
    {
        attributes.push_back(std::make_pair<std::string, std::string>(attr.name(), attr.value()));
    }

    return attributes;
}

void xml_node::add_attribute(const std::string &name, const std::string &value)
{
    d_->node.append_attribute(name.c_str()).set_value(value.c_str());
}

bool xml_node::has_attribute(const std::string &attribute_name) const
{
    return d_->node.attribute(attribute_name.c_str()) != nullptr;
}

std::string xml_node::get_attribute(const std::string &attribute_name) const
{
    return d_->node.attribute(attribute_name.c_str()).value();
}

bool xml_node::has_child(const std::string &child_name) const
{
    return d_->node.child(child_name.c_str()) != nullptr;
}

xml_node xml_node::get_child(const std::string &child_name)
{
    return xml_node(detail::xml_node_impl(d_->node.child(child_name.c_str())));
}

const xml_node xml_node::get_child(const std::string &child_name) const
{
    return xml_node(detail::xml_node_impl(d_->node.child(child_name.c_str())));
}

std::string xml_node::to_string() const
{
    return xml_serializer::serialize_node(*this);
}

} // namespace xlnt
