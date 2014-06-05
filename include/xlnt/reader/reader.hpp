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

#include <iostream>
#include <string>
#include <unordered_map>
#include <vector>

namespace xlnt {
    
class workbook;
class worksheet;
    
class reader
{
public:
    static std::unordered_map<std::string, std::pair<std::string, std::string>> read_relationships(const std::string &content);
    static std::pair<std::unordered_map<std::string, std::string>, std::unordered_map<std::string, std::string>> read_content_types(const std::string &content);
    static std::string determine_document_type(const std::unordered_map<std::string, std::pair<std::string, std::string>> &root_relationships,
                                               const std::unordered_map<std::string, std::string> &override_types);
    static worksheet read_worksheet(std::istream &handle, workbook &wb, const std::string &title, const std::vector<std::string> &string_table);
    static void read_worksheet(worksheet ws, const std::string &xml_string, const std::vector<std::string> &string_table);
    static std::vector<std::string> read_shared_string(const std::string &xml_string);
};
    
} // namespace xlnt
