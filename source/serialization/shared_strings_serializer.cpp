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
#include <xlnt/cell/text.hpp>
#include <xlnt/serialization/xml_document.hpp>
#include <xlnt/serialization/xml_node.hpp>

namespace xlnt {

xml_document shared_strings_serializer::write_shared_strings(const std::vector<text> &strings)
{
    xml_document xml;

    auto root_node = xml.add_child("sst");

    xml.add_namespace("", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

    root_node.add_attribute("uniqueCount", std::to_string(strings.size()));

    for (const auto &string : strings)
    {
        if (string.get_runs().size() == 1 && !string.get_runs().at(0).has_formatting())
        {
            root_node.add_child("si").add_child("t").set_text(string.get_plain_string());
        }
        else
        {
            for (const auto &run : string.get_runs())
            {
                auto string_item_node = root_node.add_child("si");
                auto rich_text_run_node = string_item_node.add_child("r");
                
                auto text_node = rich_text_run_node.add_child("t");
                text_node.set_text(run.get_string());
                
                if (run.has_formatting())
                {
                    auto run_properties_node = rich_text_run_node.add_child("rPr");
                    
                    run_properties_node.add_child("sz").add_attribute("val", std::to_string(run.get_size()));
                    run_properties_node.add_child("color").add_attribute("rgb", run.get_color());
                    run_properties_node.add_child("rFont").add_attribute("val", run.get_font());
                    run_properties_node.add_child("family").add_attribute("val", std::to_string(run.get_family()));
                    run_properties_node.add_child("scheme").add_attribute("val", run.get_scheme());
                }
            }
        }
    }

    return xml;
}

bool shared_strings_serializer::read_shared_strings(const xml_document &xml, std::vector<text> &strings)
{
    strings.clear();

    auto root_node = xml.get_child("sst");
    std::size_t unique_count = 0;

    if (root_node.has_attribute("uniqueCount"))
    {
        unique_count = std::stoull(root_node.get_attribute("uniqueCount"));
    }

    for (const auto &string_item_node : root_node.get_children())
    {
        if (string_item_node.get_name() != "si")
        {
            continue;
        }

        if (string_item_node.has_child("t"))
        {
            text t;
            t.set_plain_string(string_item_node.get_child("t").get_text());
            strings.push_back(t);
        }
        else if (string_item_node.has_child("r")) // possible multiple text entities.
        {
            text t;
            
            for (const auto& rich_text_run_node : string_item_node.get_children())
            {
                if (rich_text_run_node.get_name() == "r" && rich_text_run_node.has_child("t"))
                {
                    text_run run;
                    run.set_string(rich_text_run_node.get_child("t").get_text());
                    t.add_run(run);
                }
            }
            
            strings.push_back(t);
        }
    }

    if (unique_count != strings.size())
    {
        throw std::runtime_error("counts don't match");
    }

    return true;
}

} // namespace xlnt
