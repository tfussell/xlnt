// Copyright (c) 2014-2016 Thomas Fussell
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

#include "xlnt_config.hpp"

namespace xlnt {

const row_t constants::MinRow()
{
    return 1;
}
    
const row_t constants::MaxRow()
{
    if(LimitStyle == limit_style::excel)
    {
        return 1u << 20;
    }
    
    return std::numeric_limits<row_t>::max();
}

const column_t constants::MinColumn()
{
    return column_t(1);
}

const column_t constants::MaxColumn()
{
    if(LimitStyle == limit_style::excel)
    {
        return column_t(1u << 14);
    }
    
    if(LimitStyle == limit_style::openpyxl)
    {
        return column_t(18278);
    }
    
    return column_t(std::numeric_limits<column_t::index_t>::max());
}

// constants
const std::string constants::PackageProps() { return "docProps"; }
const std::string constants::PackageXl() { return "xl"; }
const std::string constants::PackageRels() { return "_rels"; }
const std::string constants::PackageTheme() { return PackageXl() + "/" + "theme"; }
const std::string constants::PackageWorksheets() { return PackageXl() + "/" + "worksheets"; }
const std::string constants::PackageDrawings() { return PackageXl() + "/" + "drawings"; }
const std::string constants::PackageCharts() { return PackageXl() + "/" + "charts"; }

const std::string constants::ArcContentTypes() { return "[Content_Types].xml"; }
const std::string constants::ArcRootRels() { return PackageRels() + "/.rels"; }
const std::string constants::ArcWorkbookRels() { return PackageXl() + "/" + PackageRels() + "/workbook.xml.rels"; }
const std::string constants::ArcCore() { return PackageProps() + "/core.xml"; }
const std::string constants::ArcApp() { return PackageProps() + "/app.xml"; }
const std::string constants::ArcWorkbook() { return PackageXl() + "/workbook.xml"; }
const std::string constants::ArcStyles() { return PackageXl() + "/styles.xml"; }
const std::string constants::ArcTheme() { return PackageTheme() + "/theme1.xml"; }
const std::string constants::ArcSharedString() { return PackageXl() + "/sharedStrings.xml"; }

const std::unordered_map<std::string, std::string> constants::Namespaces()
{
    const std::unordered_map<std::string, std::string> namespaces =
    {
        { "spreadsheetml", "http://schemas.openxmlformats.org/spreadsheetml/2006/main" },
        { "content-types", "http://schemas.openxmlformats.org/package/2006/content-types" },
        { "relationships", "http://schemas.openxmlformats.org/package/2006/relationships" },
        { "drawingml", "http://schemas.openxmlformats.org/drawingml/2006/main" },
        { "r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships" },
        { "cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties" },
        { "dc", "http://purl.org/dc/elements/1.1/" },
        { "dcterms", "http://purl.org/dc/terms/" },
        { "dcmitype", "http://purl.org/dc/dcmitype/" },
        { "xsi", "http://www.w3.org/2001/XMLSchema-instance" },
        { "vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" },
        { "xml", "http://www.w3.org/XML/1998/namespace" }
    };
    
    return namespaces;
}

const std::string constants::Namespace(const std::string &id)
{
    return Namespaces().find(id)->second;
}

}
