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
#include <vector>

#include <xlnt/xlnt_config.hpp>

namespace pugi {
class xml_document;
class xml_node;
} // namespace pugi

namespace xlnt {

class document_properties;
class manifest;
class relationship;
class worksheet;
class workbook;
class zip_file;

/// <summary>
/// Manages converting workbook to and from XML.
/// </summary>
class XLNT_CLASS workbook_serializer
{
public:
    using string_pair = std::pair<std::string, std::string>;

    workbook_serializer(workbook &wb);

    void read_workbook(const pugi::xml_document &xml);
    void read_properties_app(const pugi::xml_document &xml);
    void read_properties_core(const pugi::xml_document &xml);

    void write_workbook(pugi::xml_document &xml) const;
    void write_properties_app(pugi::xml_document &xml) const;
    void write_properties_core(pugi::xml_document &xml) const;

    void write_named_ranges(pugi::xml_node node) const;

private:
    workbook &workbook_;
};

} // namespace xlnt
