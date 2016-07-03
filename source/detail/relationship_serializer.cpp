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

#include <pugixml.hpp>
#include <sstream>

#include <detail/constants.hpp>
#include <detail/relationship_serializer.hpp>
#include <xlnt/packaging/relationship.hpp>
#include <xlnt/packaging/zip_file.hpp>

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
    pugi::xml_document xml;
    xml.load(archive_.read(make_rels_name(target)).c_str());

    auto root_node = xml.child("Relationships");

    std::vector<relationship> relationships;

    for (auto relationship_node : root_node.children("Relationship"))
    {
        std::string id = relationship_node.attribute("Id").value();
        std::string type = relationship_node.attribute("Type").value();
        std::string rel_target = relationship_node.attribute("Target").value();

        relationships.push_back(xlnt::relationship(type, id, rel_target));
    }

    return relationships;
}

bool relationship_serializer::write_relationships(const std::vector<relationship> &relationships,
                                                  const std::string &target)
{
    pugi::xml_document xml;

    auto root_node = xml.append_child("Relationships");

    root_node.append_attribute("xmlns").set_value(constants::Namespace("relationships").c_str());

    for (const auto &relationship : relationships)
    {
        auto relationship_node = root_node.append_child("Relationship");

        relationship_node.append_attribute("Id").set_value(relationship.get_id().c_str());
        relationship_node.append_attribute("Type").set_value(relationship.get_type_string().c_str());
        relationship_node.append_attribute("Target").set_value(relationship.get_target_uri().c_str());

        if (relationship.get_target_mode() == target_mode::external)
        {
            relationship_node.append_attribute("TargetMode").set_value("External");
        }
    }

    std::ostringstream ss;
    xml.save(ss);
    archive_.writestr(make_rels_name(target), ss.str());

    return true;
}
}
