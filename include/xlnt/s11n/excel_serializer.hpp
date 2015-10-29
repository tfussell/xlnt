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

#include <cstdint>
#include <ostream>
#include <string>
#include <vector>

#include <xlnt/common/zip_file.hpp>

namespace xlnt {
    
class workbook;

class excel_serializer
{
public:
    static std::string central_directory_signature();
    static std::string repair_central_directory(const std::string &original);
    
    excel_serializer(workbook &wb);
    
    bool load_workbook(const std::string &filename, bool guess_types = false, bool data_only = false);
    bool load_stream_workbook(std::istream &stream, bool guess_types = false, bool data_only = false);
    bool load_virtual_workbook(const std::vector<std::uint8_t> &bytes, bool guess_types = false, bool data_only = false);
    
    bool save_workbook(workbook &wb, const std::string &filename, bool as_template = false);
    bool save_virtual_workbook(xlnt::workbook &wb, std::vector<std::uint8_t> &bytes, bool as_template = false);
    bool save_stream_workbook(xlnt::workbook &wb, std::ostream &stream, bool as_template = false);
    
private:
    void read_data(bool guess_types, bool data_only);
    void read_shared_strings();
    void read_images();
    void read_charts();
    void read_chartsheets();
    void read_worksheets();
    void read_external_links();
    
    void write_data(bool as_template);
    void write_shared_strings();
    void write_images();
    void write_charts();
    void write_chartsheets();
    void write_worksheets();
    void write_external_links();

    workbook &wb_;
    std::vector<std::string> shared_strings_;
    zip_file archive_;
};

} // namespace xlnt
