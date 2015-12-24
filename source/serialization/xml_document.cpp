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
	d_->doc.load(xml_string.c_str());

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
