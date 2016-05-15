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

    root_node.add_attribute("count", std::to_string(strings.size()));
    root_node.add_attribute("uniqueCount", std::to_string(strings.size()));

    for (const auto &string : strings)
    {
        if (string.get_runs().size() == 1 && !string.get_runs().at(0).has_formatting())
        {
            root_node.add_child("si").add_child("t").set_text(string.get_plain_string());
        }
        else
        {
            auto string_item_node = root_node.add_child("si");
            
            for (const auto &run : string.get_runs())
            {
                auto rich_text_run_node = string_item_node.add_child("r");
                
                if (run.has_formatting())
                {
                    auto run_properties_node = rich_text_run_node.add_child("rPr");
                    
                    if (run.has_size())
                    {
                        run_properties_node.add_child("sz").add_attribute("val", std::to_string(run.get_size()));
                    }
                    
                    if (run.has_color())
                    {
                        run_properties_node.add_child("color").add_attribute("rgb", run.get_color());
                    }
                    
                    if (run.has_font())
                    {
                        run_properties_node.add_child("rFont").add_attribute("val", run.get_font());
                    }
                    
                    if (run.has_family())
                    {
                        run_properties_node.add_child("family").add_attribute("val", std::to_string(run.get_family()));
                    }
                    
                    if (run.has_scheme())
                    {
                        run_properties_node.add_child("scheme").add_attribute("val", run.get_scheme());
                    }
                }
                
                auto text_node = rich_text_run_node.add_child("t");
                text_node.set_text(run.get_string());
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
                    
                    if (rich_text_run_node.has_child("rPr"))
                    {
                        auto run_properties_node = rich_text_run_node.get_child("rPr");
                        
                        if (run_properties_node.has_child("sz"))
                        {
                            run.set_size(std::stoull(run_properties_node.get_child("sz").get_attribute("val")));
                        }
                        
                        if (run_properties_node.has_child("rFont"))
                        {
                            run.set_font(run_properties_node.get_child("rFont").get_attribute("val"));
                        }
                        
                        if (run_properties_node.has_child("color"))
                        {
                            run.set_color(run_properties_node.get_child("color").get_attribute("rgb"));
                        }
                        
                        if (run_properties_node.has_child("family"))
                        {
                            run.set_family(std::stoull(run_properties_node.get_child("family").get_attribute("val")));
                        }
                        
                        if (run_properties_node.has_child("scheme"))
                        {
                            run.set_scheme(run_properties_node.get_child("scheme").get_attribute("val"));
                        }
                    }
                    
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
