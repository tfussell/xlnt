#pragma once

#include <detail/include_pugixml.hpp>

namespace xlnt {
namespace detail {

struct xml_node_impl
{
    xml_node_impl()
    {
    }
    
    explicit xml_node_impl(pugi::xml_node n) : node(n)
    {
    }
    
    pugi::xml_node node;
};

} // namespace detail
} // namespace xlnt
