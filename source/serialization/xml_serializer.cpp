// Copyright (c) 2014-2016 Thomas Fussell
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, WRISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE
//
// @license: http://www.opensource.org/licenses/mit-license.php
// @author: see AUTHORS file
#include <sstream>

#include <xlnt/serialization/xml_serializer.hpp>
#include <xlnt/serialization/xml_document.hpp>
#include <xlnt/serialization/xml_node.hpp>

#include "detail/include_pugixml.hpp"
#include "detail/xml_document_impl.hpp"
#include "detail/xml_node_impl.hpp"

namespace xlnt {

std::string xml_serializer::serialize(const xml_document &xml)
{
    std::ostringstream ss;
    xml.d_->doc.save(ss, "   ", pugi::format_default, pugi::encoding_utf8);

    return ss.str();
}

xml_document xml_serializer::deserialize(const std::string &xml_string)
{
    xml_document doc;
    doc.d_->doc.load(xml_string.c_str());

    return doc;
}

std::string xml_serializer::serialize_node(const xml_node &xml)
{
    pugi::xml_document doc;
    doc.append_copy(xml.d_->node);

    std::ostringstream ss;
    doc.save(ss);

    return ss.str();
}

} // namespace xlnt
