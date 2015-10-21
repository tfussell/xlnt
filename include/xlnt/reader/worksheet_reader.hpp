// Copyright (c) 2015 Thomas Fussell
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

#include <iostream>
#include <string>
#include <unordered_map>
#include <vector>

namespace xlnt {
    
class workbook;
class worksheet;

worksheet read_worksheet(std::istream &handle, workbook &wb, const std::string &title, const std::vector<std::string> &string_table, const std::unordered_map<int, std::string> &custom_number_formats);
void read_worksheet(worksheet ws, std::istream &stream, const std::vector<std::string> &string_table, const std::vector<int> &number_format_ids, const std::unordered_map<int, std::string> &custom_number_formats);
void read_worksheet(worksheet ws, const std::string &xml_string, const std::vector<std::string> &string_table, const std::vector<int> &number_format_ids, const std::unordered_map<int, std::string> &custom_number_formats);
std::string read_dimension(const std::string &xml_string);

} // namespace xlnt
