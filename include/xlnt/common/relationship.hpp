// Copyright (c) 2014 Thomas Fussell
// Copyright (c) 2010-2014 openpyxl
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

#include <string>

namespace xlnt {
    
/// <summary>
/// Specifies whether the target of a relationship is inside or outside the Package.
/// </summary>
enum class target_mode
{
    /// <summary>
    /// The relationship references a part that is inside the package.
    /// </summary>
    external,
    /// <summary>
    /// The relationship references a resource that is external to the package.
    /// </summary>
    internal
};

/// <summary>
/// Represents an association between a source Package or part, and a target object which can be a part or external resource.
/// </summary>
class relationship
{
public:
    enum class type
    {
        invalid,
        hyperlink,
        drawing,
        worksheet,
        chartsheet,
        shared_strings,
        styles,
        theme,
        extended_properties,
        core_properties,
        office_document,
        custom_xml
    };
    
    static type type_from_string(const std::string &type_string)
    {
        if(type_string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties")
        {
            return type::extended_properties;
        }
        else if(type_string == "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties")
        {
            return type::core_properties;
        }
        else if(type_string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument")
        {
            return type::office_document;
        }
        else if(type_string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")
        {
            return type::worksheet;
        }
        else if(type_string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings")
        {
            return type::shared_strings;
        }
        else if(type_string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles")
        {
            return type::styles;
        }
        else if(type_string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
        {
            return type::theme;
        }
        else if(type_string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink")
        {
            return type::hyperlink;
        }
        else if(type_string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet")
        {
            return type::chartsheet;
        }
        else if(type_string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml")
        {
            return type::custom_xml;
        }
        
        return type::invalid;
    }
    
    static std::string type_to_string(type t)
    {
        switch(t)
        {
            case type::extended_properties: return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
            case type::core_properties: return "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
            case type::office_document: return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
            case type::worksheet: return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
            case type::shared_strings: return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
            case type::styles: return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
            case type::theme: return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
            case type::hyperlink: return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
            case type::chartsheet: return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
            case type::custom_xml: return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
            default: return "??";
        }
    }
    
    relationship();
    relationship(const std::string &t, const std::string &r_id = "", const std::string &target_uri = "") : relationship(type_from_string(t), r_id, target_uri) {}
    relationship(type t, const std::string &r_id = "", const std::string &target_uri = "");
    
    /// <summary>
    /// gets a string that identifies the relationship.
    /// </summary>
    std::string get_id() const { return id_; }
    
    /// <summary>
    /// gets the URI of the package or part that owns the relationship.
    /// </summary>
    std::string get_source_uri() const { return source_uri_; }
    
    /// <summary>
    /// gets a value that indicates whether the target of the relationship is or External to the Package.
    /// </summary>
    target_mode get_target_mode() const { return target_mode_; }
    
    /// <summary>
    /// gets the URI of the target resource of the relationship.
    /// </summary>
    std::string get_target_uri() const { return target_uri_; }
    
    type get_type() const { return type_; }
    std::string get_type_string() const { return type_to_string(type_); }

    friend bool operator==(const relationship &left, const relationship &right)
    {
        return left.type_ == right.type_
            && left.id_ == right.id_
            && left.source_uri_ == right.source_uri_
            && left.target_uri_ == right.target_uri_
            && left.target_mode_ == right.target_mode_;
    }
    
private:
    type type_;
    std::string id_;
    std::string source_uri_;
    std::string target_uri_;
    target_mode target_mode_;
};
    
} // namespace xlnt
