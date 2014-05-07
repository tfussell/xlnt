/*
# Copyright(c) 2010 - 2011 openpyxl
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files(the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and / or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions :
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.
*/
#pragma once

#include <string>
#include <unordered_map>

namespace xlnt {

// Constants for fixed paths in a file and xml namespace urls.

const int MIN_ROW = 0;
const int MIN_COLUMN = 0;
const int MAX_COLUMN = 16384;
const int MAX_ROW = 1048576;

// constants
const std::string PACKAGE_PROPS = "docProps";
const std::string PACKAGE_XL = "xl";
const std::string PACKAGE_RELS = "_rels";
const std::string PACKAGE_THEME = PACKAGE_XL + "/" + "theme";
const std::string PACKAGE_WORKSHEETS = PACKAGE_XL + "/" + "worksheets";
const std::string PACKAGE_DRAWINGS = PACKAGE_XL + "/" + "drawings";
const std::string PACKAGE_CHARTS = PACKAGE_XL + "/" + "charts";

const std::string ARC_CONTENT_TYPES = "[Content_Types].xml";
const std::string ARC_ROOT_RELS = PACKAGE_RELS + "/.rels";
const std::string ARC_WORKBOOK_RELS = PACKAGE_XL + "/" + PACKAGE_RELS + "/workbook.xml.rels";
const std::string ARC_CORE = PACKAGE_PROPS + "/core.xml";
const std::string ARC_APP = PACKAGE_PROPS + "/app.xml";
const std::string ARC_WORKBOOK = PACKAGE_XL + "/workbook.xml";
const std::string ARC_STYLE = PACKAGE_XL + "/styles.xml";
const std::string ARC_THEME = PACKAGE_THEME + "/theme1.xml";
const std::string ARC_SHARED_STRINGS = PACKAGE_XL + "/sharedStrings.xml";

std::unordered_map<std::string, std::string> NAMESPACES = {
    {"cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"},
    {"dc", "http://purl.org/dc/elements/1.1/"},
    {"dcterms", "http://purl.org/dc/terms/"},
    {"dcmitype", "http://purl.org/dc/dcmitype/"},
    {"xsi", "http://www.w3.org/2001/XMLSchema-instance"},
    {"vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"},
    {"xml", "http://www.w3.org/XML/1998/namespace"}
};

} // namespace xlnt
