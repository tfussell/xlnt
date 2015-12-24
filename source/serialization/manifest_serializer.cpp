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
#include <xlnt/packaging/manifest.hpp>
#include <xlnt/serialization/manifest_serializer.hpp>
#include <xlnt/serialization/xml_document.hpp>
#include <xlnt/serialization/xml_node.hpp>

#include <detail/constants.hpp>

namespace xlnt {

manifest_serializer::manifest_serializer(manifest &m) : manifest_(m)
{
}

void manifest_serializer::read_manifest(const xml_document &xml)
{
    const auto root_node = xml.get_child("Types");

    for (const auto child : root_node.get_children())
    {
        if (child.get_name() == "Default")
        {
            manifest_.add_default_type(child.get_attribute("Extension"), child.get_attribute("ContentType"));
        }
        else if (child.get_name() == "Override")
        {
            manifest_.add_override_type(child.get_attribute("PartName"), child.get_attribute("ContentType"));
        }
    }
}

xml_document manifest_serializer::write_manifest() const
{
    xml_document xml;

    auto root_node = xml.add_child("Types");
    xml.add_namespace("", constants::Namespace("content-types"));

    for (const auto default_type : manifest_.get_default_types())
    {
        auto type_node = root_node.add_child("Default");
        type_node.add_attribute("Extension", default_type.get_extension());
        type_node.add_attribute("ContentType", default_type.get_content_type());
    }

    for (const auto override_type : manifest_.get_override_types())
    {
        auto type_node = root_node.add_child("Override");
        type_node.add_attribute("PartName", override_type.get_part_name());
        type_node.add_attribute("ContentType", override_type.get_content_type());
    }

    return xml;
}

std::string manifest_serializer::determine_document_type() const
{
    if (!manifest_.has_override_type(constants::ArcWorkbook()))
    {
        return "unsupported";
    }

    std::string type = manifest_.get_override_type(constants::ArcWorkbook());

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

    return "unsupported";
}

} // namespace xlnt
