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
    
class document_properties;
class relationship;
class style;
class workbook;
class worksheet;
class zip_file;

class reader
{
public:
    static const std::string CentralDirectorySignature;
    static std::string repair_central_directory(const std::string &original);
    static void fast_parse(worksheet ws, std::istream &xml_source, const std::vector<std::string> &shared_string, const std::vector<style> &style_table, std::size_t color_index);
    static std::vector<relationship> read_relationships(const zip_file &content, const std::string &filename);
    static std::vector<std::pair<std::string, std::string>> read_content_types(const zip_file &archive);
    static std::string determine_document_type(const std::vector<std::pair<std::string, std::string>> &override_types);
    static worksheet read_worksheet(std::istream &handle, workbook &wb, const std::string &title, const std::vector<std::string> &string_table);
    static void read_worksheet(worksheet ws, const std::string &xml_string, const std::vector<std::string> &string_table, const std::vector<int> &number_format_ids);
    static std::vector<std::string> read_shared_string(const std::string &xml_string);
    static std::string read_dimension(const std::string &xml_string);
    static document_properties read_properties_core(const std::string &xml_string);
    static std::vector<std::pair<std::string,std::string>> read_sheets(const zip_file &archive);
    static workbook load_workbook(const std::string &filename, bool guess_types = false, bool data_only = false);
    static std::vector<std::pair<std::string, std::string>> detect_worksheets(const zip_file &archive);
};
    
} // namespace xlnt
