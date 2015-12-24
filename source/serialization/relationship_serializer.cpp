// Copyright (c) 2014-2016 Thomas Fussell
// Copyright (c) 2010-2015 openpyxl
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
#include <xlnt/serialization/relationship_serializer.hpp>

#include <xlnt/packaging/relationship.hpp>
#include <xlnt/packaging/zip_file.hpp>
#include <xlnt/serialization/xml_document.hpp>
#include <xlnt/serialization/xml_node.hpp>

#include "detail/constants.hpp"

namespace {

std::string make_rels_name(const std::string &target)
{
    const char sep = '/';

    if (target.empty() || target.back() == sep)
    {
        return target + "_rels/.rels";
    }

    auto sep_pos = target.find_last_of(sep);
    auto first_part = target.substr(0, sep_pos + 1);
    auto last_part = target.substr(sep_pos + 1);

    return first_part + "_rels/" + last_part + ".rels";
}
}

namespace xlnt {

relationship_serializer::relationship_serializer(zip_file &archive) : archive_(archive)
{
}

std::vector<relationship> relationship_serializer::read_relationships(const std::string &target)
{
    xml_document xml;
    xml.from_string(archive_.read(make_rels_name(target)));

    auto root_node = xml.get_child("Relationships");

    std::vector<relationship> relationships;

    for (auto relationship_node : root_node.get_children())
    {
        if (relationship_node.get_name() != "Relationship")
        {
            continue;
        }

        std::string id = relationship_node.get_attribute("Id");
        std::string type = relationship_node.get_attribute("Type");
        std::string rel_target = relationship_node.get_attribute("Target");

        relationships.push_back(xlnt::relationship(type, id, rel_target));
    }

    return relationships;
}

bool relationship_serializer::write_relationships(const std::vector<relationship> &relationships,
                                                  const std::string &target)
{
    xml_document xml;

    auto root_node = xml.add_child("Relationships");

    xml.add_namespace("", constants::Namespace("relationships"));

    for (const auto &relationship : relationships)
    {
        auto relationship_node = root_node.add_child("Relationship");

        relationship_node.add_attribute("Id", relationship.get_id());
        relationship_node.add_attribute("Target", relationship.get_target_uri());
        relationship_node.add_attribute("Type", relationship.get_type_string());

        if (relationship.get_target_mode() == target_mode::external)
        {
            relationship_node.add_attribute("TargetMode", "External");
        }
    }

    archive_.writestr(make_rels_name(target), xml.to_string());

    return true;
}
}
