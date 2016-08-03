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

#include <detail/constants.hpp>
#include <detail/manifest_serializer.hpp>
#include <xlnt/packaging/manifest.hpp>

namespace xlnt {

manifest_serializer::manifest_serializer(manifest &m) : manifest_(m)
{
}

void manifest_serializer::read_manifest(const pugi::xml_document &xml)
{
    const auto root_node = xml.child("Types");

    for (const auto child : root_node.children())
    {
        if (child.name() == std::string("Default"))
        {
            manifest_.add_default_type(child.attribute("Extension").value(), child.attribute("ContentType").value());
        }
        else if (child.name() == std::string("Override"))
        {
            manifest_.add_override_type(path(child.attribute("PartName").value()), child.attribute("ContentType").value());
        }
    }
}

void manifest_serializer::write_manifest(pugi::xml_document &xml) const
{
    auto root_node = xml.append_child("Types");
    root_node.append_attribute("xmlns").set_value(constants::get_namespace("content-types").c_str());

    for (const auto default_type : manifest_.get_default_types())
    {
        auto type_node = root_node.append_child("Default");
        type_node.append_attribute("Extension").set_value(default_type.second.get_extension().c_str());
        type_node.append_attribute("ContentType").set_value(default_type.second.get_content_type().c_str());
    }

    for (const auto override_type : manifest_.get_override_types())
    {
        auto type_node = root_node.append_child("Override");
        type_node.append_attribute("PartName").set_value(override_type.second.get_part().to_string('/').c_str());
        type_node.append_attribute("ContentType").set_value(override_type.second.get_content_type().c_str());
    }
}

std::string manifest_serializer::determine_document_type() const
{
    for (auto current_override_type : manifest_.get_override_types())
    {
        auto type = current_override_type.second.get_content_type();

        if (type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
        {
            return "excel";
        }
        else if (type == "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml")
        {
            return "powerpoint";
        }
        else if (type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")
        {
            return "word";
        }
    }

    return "unsupported";
}

} // namespace xlnt
