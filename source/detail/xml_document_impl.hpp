#pragma once

#include <string>

#include <detail/include_pugixml.hpp>

namespace xlnt {
namespace detail {

struct xml_document_impl
{
    std::string encoding;
    pugi::xml_document doc;
};
    
} // namespace detail
} // namespace xlnt
