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
#pragma once

#include <memory>
#include <string>
#include <vector>

#include <xlnt/xlnt_config.hpp>

namespace xlnt {
namespace detail { struct xml_document_impl; }

class xml_node;
class xml_serializer;

/// <summary>
/// Abstracts an XML document from a particular implementation.
/// </summary>
class XLNT_CLASS xml_document
{
  public:
    using string_pair = std::pair<std::string, std::string>;

    xml_document();
    xml_document(const xml_document &other);
    xml_document(xml_document &&other);
    ~xml_document();

    xml_document &operator=(const xml_document &other);
    xml_document &operator=(xml_document &&other);

    void set_encoding(const std::string &encoding);
    void add_namespace(const std::string &id, const std::string &uri);

    xml_node add_child(const xml_node &child);
    xml_node add_child(const std::string &child_name);

    xml_node get_root();
    const xml_node get_root() const;

    xml_node get_child(const std::string &child_name);
    const xml_node get_child(const std::string &child_name) const;

    std::string to_string() const;
    xml_document &from_string(const std::string &xml_string);

  private:
    friend class xml_serializer;
    std::unique_ptr<detail::xml_document_impl> d_;
};

} // namespace xlnt
