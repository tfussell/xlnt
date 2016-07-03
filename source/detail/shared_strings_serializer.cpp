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

#include <detail/shared_strings_serializer.hpp>
#include <xlnt/cell/text.hpp>

namespace {

std::size_t string_to_size_t(const std::string &s)
{
#if ULLONG_MAX == SIZE_MAX
	return std::stoull(s);
#else
	return std::stoul(s);
#endif
} // namespace

}

namespace xlnt {

void shared_strings_serializer::write_shared_strings(const std::vector<text> &strings, pugi::xml_document &xml)
{
    auto root_node = xml.append_child("sst");

    root_node.append_attribute("xmlns").set_value("http://schemas.openxmlformats.org/spreadsheetml/2006/main");

    root_node.append_attribute("count").set_value(std::to_string(strings.size()).c_str());
    root_node.append_attribute("uniqueCount").set_value(std::to_string(strings.size()).c_str());

    for (const auto &string : strings)
    {
        if (string.get_runs().size() == 1 && !string.get_runs().at(0).has_formatting())
        {
            root_node.append_child("si").append_child("t").text().set(string.get_plain_string().c_str());
        }
        else
        {
            auto string_item_node = root_node.append_child("si");
            
            for (const auto &run : string.get_runs())
            {
                auto rich_text_run_node = string_item_node.append_child("r");
                
                if (run.has_formatting())
                {
                    auto run_properties_node = rich_text_run_node.append_child("rPr");
                    
                    if (run.has_size())
                    {
                        run_properties_node.append_child("sz").append_attribute("val").set_value(std::to_string(run.get_size()).c_str());
                    }
                    
                    if (run.has_color())
                    {
                        run_properties_node.append_child("color").append_attribute("rgb").set_value(run.get_color().c_str());
                    }
                    
                    if (run.has_font())
                    {
                        run_properties_node.append_child("rFont").append_attribute("val").set_value(run.get_font().c_str());
                    }
                    
                    if (run.has_family())
                    {
                        run_properties_node.append_child("family").append_attribute("val").set_value(std::to_string(run.get_family()).c_str());
                    }
                    
                    if (run.has_scheme())
                    {
                        run_properties_node.append_child("scheme").append_attribute("val").set_value(run.get_scheme().c_str());
                    }
                }
                
                auto text_node = rich_text_run_node.append_child("t");
                text_node.text().set(run.get_string().c_str());
            }
        }
    }
}

bool shared_strings_serializer::read_shared_strings(const pugi::xml_document &xml, std::vector<text> &strings)
{
    strings.clear();

    auto root_node = xml.child("sst");
    std::size_t unique_count = 0;

    if (root_node.attribute("uniqueCount"))
    {
        unique_count = string_to_size_t(root_node.attribute("uniqueCount").value());
    }

    for (const auto string_item_node : root_node.children("si"))
    {
        if (string_item_node.child("t"))
        {
            text t;
            t.set_plain_string(string_item_node.child("t").text().get());
            strings.push_back(t);
        }
        else if (string_item_node.child("r")) // possible multiple text entities.
        {
            text t;
            
            for (const auto& rich_text_run_node : string_item_node.children("r"))
            {
                if (rich_text_run_node.child("t"))
                {
                    text_run run;
                    
                    run.set_string(rich_text_run_node.child("t").text().get());
                    
                    if (rich_text_run_node.child("rPr"))
                    {
                        auto run_properties_node = rich_text_run_node.child("rPr");
                        
                        if (run_properties_node.child("sz"))
                        {
                            run.set_size(string_to_size_t(run_properties_node.child("sz").attribute("val").value()));
                        }
                        
                        if (run_properties_node.child("rFont"))
                        {
                            run.set_font(run_properties_node.child("rFont").attribute("val").value());
                        }
                        
                        if (run_properties_node.child("color"))
                        {
                            run.set_color(run_properties_node.child("color").attribute("rgb").value());
                        }
                        
                        if (run_properties_node.child("family"))
                        {
                            run.set_family(string_to_size_t(run_properties_node.child("family").attribute("val").value()));
                        }
                        
                        if (run_properties_node.child("scheme"))
                        {
                            run.set_scheme(run_properties_node.child("scheme").attribute("val").value());
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
