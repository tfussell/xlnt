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
#include <xlnt/packaging/relationship.hpp>

namespace xlnt {

relationship::relationship(type t, const std::string &r_id, const std::string &target_uri)
    : type_(t), id_(r_id), source_uri_(""), target_uri_(target_uri), target_mode_(target_mode::internal)
{
    if (t == type::hyperlink)
    {
        target_mode_ = target_mode::external;
    }
}

relationship::relationship()
    : type_(type::invalid), id_(""), source_uri_(""), target_uri_(""), target_mode_(target_mode::internal)
{
}

relationship::type relationship::type_from_string(const std::string &type_string)
{
    if (type_string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties")
    {
        return type::extended_properties;
    }
    else if (type_string == "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties")
    {
        return type::core_properties;
    }
    else if (type_string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument")
    {
        return type::office_document;
    }
    else if (type_string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")
    {
        return type::worksheet;
    }
    else if (type_string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings")
    {
        return type::shared_strings;
    }
    else if (type_string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles")
    {
        return type::styles;
    }
    else if (type_string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
    {
        return type::theme;
    }
    else if (type_string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink")
    {
        return type::hyperlink;
    }
    else if (type_string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet")
    {
        return type::chartsheet;
    }
    else if (type_string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml")
    {
        return type::custom_xml;
    }

    return type::invalid;
}

std::string relationship::type_to_string(type t)
{
    switch (t)
    {
    case type::extended_properties:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
    case type::core_properties:
        return "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
    case type::office_document:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    case type::worksheet:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
    case type::shared_strings:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
    case type::styles:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    case type::theme:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
    case type::hyperlink:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    case type::chartsheet:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
    case type::custom_xml:
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
    default:
        return "??";
    }
}

relationship::relationship(const std::string &t, const std::string &r_id, const std::string &target_uri)
    : relationship(type_from_string(t), r_id, target_uri)
{
}

std::string relationship::get_id() const
{
    return id_;
}

std::string relationship::get_source_uri() const
{
    return source_uri_;
}

target_mode relationship::get_target_mode() const
{
    return target_mode_;
}

std::string relationship::get_target_uri() const
{
    return target_uri_;
}

relationship::type relationship::get_type() const
{
    return type_;
}
std::string relationship::get_type_string() const
{
    return type_to_string(type_);
}

bool relationship::operator==(const relationship &rhs) const
{
    return type_ == rhs.type_ && id_ == rhs.id_ && source_uri_ == rhs.source_uri_ &&
           target_uri_ == rhs.target_uri_ && target_mode_ == rhs.target_mode_;
}

} // namespace xlnt
