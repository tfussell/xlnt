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
#pragma once

#include <string>

#include <xlnt/xlnt_config.hpp>

namespace xlnt {

/// <summary>
/// Specifies whether the target of a relationship is inside or outside the Package.
/// </summary>
enum class XLNT_CLASS target_mode
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
/// Represents an association between a source Package or part, and a target object which can be a part or external
/// resource.
/// </summary>
class XLNT_CLASS relationship
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

    static type type_from_string(const std::string &type_string);

    static std::string type_to_string(type t);
    
    relationship();
    
    relationship(const std::string &t, const std::string &r_id = "", const std::string &target_uri = "");
    
    relationship(type t, const std::string &r_id = "", const std::string &target_uri = "");

    /// <summary>
    /// gets a string that identifies the relationship.
    /// </summary>
    std::string get_id() const;

    /// <summary>
    /// gets the URI of the package or part that owns the relationship.
    /// </summary>
    std::string get_source_uri() const;

    /// <summary>
    /// gets a value that indicates whether the target of the relationship is or External to the Package.
    /// </summary>
    target_mode get_target_mode() const;

    /// <summary>
    /// gets the URI of the target resource of the relationship.
    /// </summary>
    std::string get_target_uri() const;

    type get_type() const;
    
    std::string get_type_string() const;
    
    bool operator==(const relationship &rhs) const;

  private:
    type type_;
    std::string id_;
    std::string source_uri_;
    std::string target_uri_;
    target_mode target_mode_;
};

} // namespace xlnt
