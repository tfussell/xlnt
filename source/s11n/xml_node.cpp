#include <xlnt/s11n/xml_node.hpp>

namespace xlnt {

xml_node::xml_node()
{
}

xml_node::xml_node(const std::string &name)
{
    set_name(name);
}

std::string xml_node::get_name() const
{
    return name_;
}

void xml_node::set_name(const std::string &name)
{
    name_ = name;
}

bool xml_node::has_text() const
{
    return has_text_;
}

std::string xml_node::get_text() const
{
    return text_;
}

void xml_node::set_text(const std::string &text)
{
    text_ = text;
    has_text_ = true;
}

const std::vector<xml_node> &xml_node::get_children() const
{
    return children_;
}

xml_node &xml_node::add_child(const xml_node &child)
{
    has_text_ = false;
    text_.clear();
    
    children_.push_back(child);
    return children_.back();
}

xml_node &xml_node::add_child(const std::string &child_name)
{
    return add_child(xml_node(child_name));
}

const std::vector<string_pair> &xml_node::get_attributes() const
{
    return attributes_;
}

void xml_node::add_attribute(const std::string &name, const std::string &value)
{
    attributes_.push_back(std::make_pair(name, value));
}

bool xml_node::has_attribute(const std::string &attribute_name) const
{
    return std::find(attributes_.begin(), attributes_.end(),
        [&](const string_pair &p) { return p.first == attribute_name; }) != attributes_.end();
}

std::string xml_node::get_attribute(const std::string &attribute_name) const
{
    auto match = std::find(attributes_.begin(), attributes_.end(),
        [&](const string_pair &p) { return p.first == attribute_name; });
    
    if(match == attributes_.end())
    {
        throw std::runtime_error("attribute doesn't exist: " + attribute_name);
    }
    
    return match->second;
}

bool xml_node::has_child(const std::string &child_name) const
{
    return std::find(children_.begin(), children_.end(),
        [&](const xml_node &n) { return n.get_name() == child_name; }) != children_.end();
}

const xml_node &xml_node::get_child(const std::string &child_name) const
{
    auto match = std::find(children_.begin(), children_.end(),
        [&](const xml_node &n) { return n.get_name() == child_name; });
    
    if(match == attributes_.end())
    {
        throw std::runtime_error("child doesn't exist: " + child_name);
    }
    
    return *match;
}

} // namespace xlnt
