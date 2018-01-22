// Copyright (c) 2014-2018 Thomas Fussell
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

#include <limits>

#include <detail/constants.hpp>
#include <xlnt/xlnt_config.hpp>
#include <xlnt/utils/exceptions.hpp>

namespace xlnt {

row_t constants::min_row()
{
    return 1;
}

row_t constants::max_row()
{
    return std::numeric_limits<row_t>::max();
}

const column_t constants::min_column()
{
    return column_t(1);
}

const column_t constants::max_column()
{
    return column_t(std::numeric_limits<column_t::index_t>::max());
}

// constants
const path constants::package_properties()
{
    return path("docProps");
}

const path constants::package_xl()
{
    return path("/xl");
}

const path constants::package_root_rels()
{
    return path(std::string("_rels"));
}

const path constants::package_theme()
{
    return package_xl().append("theme");
}

const path constants::package_worksheets()
{
    return package_xl().append("worksheets");
}

const path constants::package_drawings()
{
    return package_xl().append("drawings");
}

const path constants::part_content_types()
{
    return path("[Content_Types].xml");
}

const path constants::part_root_relationships()
{
    return package_root_rels().append(".rels");
}

const path constants::part_core()
{
    return package_properties().append("core.xml");
}

const path constants::part_app()
{
    return package_properties().append("app.xml");
}

const path constants::part_workbook()
{
    return package_xl().append("workbook.xml");
}

const path constants::part_styles()
{
    return package_xl().append("styles.xml");
}

const path constants::part_theme()
{
    return package_theme().append("theme1.xml");
}

const path constants::part_shared_strings()
{
    return package_xl().append("sharedStrings.xml");
}

const std::unordered_map<std::string, std::string> &constants::namespaces()
{
    static const std::unordered_map<std::string, std::string> *namespaces =
        new std::unordered_map<std::string, std::string>{
            {"spreadsheetml", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"},
            {"content-types", "http://schemas.openxmlformats.org/package/2006/content-types"},
            {"relationships", "http://schemas.openxmlformats.org/package/2006/relationships"},
            {"drawingml", "http://schemas.openxmlformats.org/drawingml/2006/main"},
            {"workbook", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"},
            {"core-properties", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"},
            {"extended-properties", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"},
            {"custom-properties", "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"},

            {"encryption", "http://schemas.microsoft.com/office/2006/encryption"},
            {"encryption-password", "http://schemas.microsoft.com/office/2006/keyEncryptor/password"},
            {"encryption-certificate", "http://schemas.microsoft.com/office/2006/keyEncryptor/certificate"},

            {"dc", "http://purl.org/dc/elements/1.1/"},
            {"dcterms", "http://purl.org/dc/terms/"},
            {"dcmitype", "http://purl.org/dc/dcmitype/"},
            {"mc", "http://schemas.openxmlformats.org/markup-compatibility/2006"},
            {"mx", "http://schemas.microsoft.com/office/mac/excel/2008/main"},
            {"r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"},
            {"thm15", "http://schemas.microsoft.com/office/thememl/2012/main"},
            {"vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"},
            {"x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"},
            {"x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"},
            {"x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"},
            {"x15ac", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"},
            {"xml", "http://www.w3.org/XML/1998/namespace"},
            {"xsi", "http://www.w3.org/2001/XMLSchema-instance"},

            {"loext", "http://schemas.libreoffice.org/"}
        };

    return *namespaces;
}

const std::string &constants::ns(const std::string &id)
{
    auto match = namespaces().find(id);

    if (match == namespaces().end())
    {
        throw xlnt::exception("bad namespace");
    }

    return match->second;
}

} // namespace xlnt
