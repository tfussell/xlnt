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
#include <xlnt/serialization/shared_strings_serializer.hpp>
#include <xlnt/serialization/xml_document.hpp>
#include <xlnt/serialization/xml_node.hpp>

namespace xlnt {

xml_document shared_strings_serializer::write_shared_strings(const std::vector<std::string> &strings)
{
    xml_document xml;

    auto root_node = xml.add_child("sst");

    xml.add_namespace("", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

    root_node.add_attribute("uniqueCount", std::to_string(strings.size()));

    for (const auto &string : strings)
    {
        root_node.add_child("si").add_child("t").set_text(string);
    }

    return xml;
}

bool shared_strings_serializer::read_shared_strings(const xml_document &xml, std::vector<std::string> &strings)
{
    strings.clear();

    auto root_node = xml.get_child("sst");
    auto unique_count = 0;

    if (root_node.has_attribute("uniqueCount"))
    {
        unique_count = std::stoull(root_node.get_attribute("uniqueCount"));
    }

    for (const auto &si_node : root_node.get_children())
    {
        if (si_node.get_name() != "si")
        {
            continue;
        }

        if (si_node.has_child("t"))
        {
            strings.push_back(si_node.get_child("t").get_text());
        }
        else if (si_node.has_child("r"))
        {
            strings.push_back(si_node.get_child("r").get_child("t").get_text());
        }
    }

    if (unique_count != strings.size())
    {
        throw std::runtime_error("counts don't match");
    }

    return true;
}

} // namespace xlnt
