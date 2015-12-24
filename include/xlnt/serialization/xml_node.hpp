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
namespace detail { struct xml_node_impl; }

class xml_document;

/// <summary>
/// Abstracts an XML node from a particular implementation.
/// </summary>
class XLNT_CLASS xml_node
{
  public:
    using string_pair = std::pair<std::string, std::string>;

    xml_node();
    xml_node(const xml_node &other);
    ~xml_node();

    xml_node &operator=(const xml_node &other);

    std::string get_name() const;
    void set_name(const std::string &name);

    bool has_text() const;
    std::string get_text() const;
    void set_text(const std::string &text);

    const std::vector<xml_node> get_children() const;
    bool has_child(const std::string &child_name) const;
    xml_node get_child(const std::string &child_name);
    const xml_node get_child(const std::string &child_name) const;
    xml_node add_child(const xml_node &child);
    xml_node add_child(const std::string &child_name);

    const std::vector<string_pair> get_attributes() const;
    bool has_attribute(const std::string &attribute_name) const;
    std::string get_attribute(const std::string &attribute_name) const;
    void add_attribute(const std::string &name, const std::string &value);

    std::string to_string() const;

  private:
    friend class xml_document;
    friend class xml_serializer;
    xml_node(const detail::xml_node_impl &d);
    std::unique_ptr<detail::xml_node_impl> d_;
};

} // namespace xlnt
