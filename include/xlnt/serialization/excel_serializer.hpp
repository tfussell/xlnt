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

#include <cstdint>
#include <ostream>
#include <string>
#include <vector>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/packaging/zip_file.hpp>

namespace xlnt {

class workbook;

/// <summary>
/// Handles reading and writing workbooks from an actual XLSX archive
/// using other serializers.
/// </summary>
class XLNT_CLASS excel_serializer
{
  public:
    /// <summary>
    ///
    /// </summary>
    static const std::string central_directory_signature();

    /// <summary>
    ///
    /// </summary>
    static std::string repair_central_directory(const std::string &original);

    /// <summary>
    /// Construct an excel_serializer which operates on wb.
    /// </summary>
    excel_serializer(workbook &wb);

    /// <summary>
    /// Create a ZIP file in memory, load archive from filename, then populate workbook
    /// with data from archive.
    /// </summary>
    bool load_workbook(const std::string &filename, bool guess_types = false, bool data_only = false);

    /// <summary>
    /// Create a ZIP file in memory, load archive from stream, then populate workbook
    /// with data from archive.
    /// </summary>
    bool load_stream_workbook(std::istream &stream, bool guess_types = false, bool data_only = false);

    /// <summary>
    /// Create a ZIP file in memory, load archive from bytes, then populate workbook
    /// with data from archive.
    /// </summary>
    bool load_virtual_workbook(const std::vector<std::uint8_t> &bytes, bool guess_types = false,
                               bool data_only = false);

    /// <summary>
    /// Create a ZIP file in memory, save workbook to this archive, then save archive
    /// to filename.
    /// </summary>
    bool save_workbook(const std::string &filename, bool as_template = false);

    /// <summary>
    /// Create a ZIP file in memory, save workbook to this archive, then assign ZIP file
    /// binary data to bytes.
    /// </summary>
    bool save_virtual_workbook(std::vector<std::uint8_t> &bytes, bool as_template = false);

    /// <summary>
    /// Create a ZIP file in memory, save workbook to this archive, then copy ZIP file
    /// binary data to stream.
    /// </summary>
    bool save_stream_workbook(std::ostream &stream, bool as_template = false);

  private:
    /// <summary>
    /// Reads all files in archive and populates workbook with associated data
    /// using other appropriate serializers such as workbook_serializer.
    /// </summary>
    void read_data(bool guess_types, bool data_only);

    /// <summary>
    /// Read xl/sharedStrings.xml from internal archive and add shared strings to workbook.
    /// </summary>
    void read_shared_strings();

    /// <summary>
    ///
    /// </summary>
    void read_images();

    /// <summary>
    ///
    /// </summary>
    void read_charts();

    /// <summary>
    ///
    /// </summary>
    void read_chartsheets();

    /// <summary>
    ///
    /// </summary>
    void read_worksheets();

    /// <summary>
    ///
    /// </summary>
    void read_external_links();

    /// <summary>
    /// Write all files needed to create a valid XLSX file which represents all
    /// data contained in workbook.
    /// </summary>
    void write_data(bool as_template);

    /// <summary>
    /// Write xl/sharedStrings.xml to internal archive based on shared strings in workbook.
    /// </summary>
    void write_shared_strings();

    /// <summary>
    ///
    /// </summary>
    void write_images();

    /// <summary>
    ///
    /// </summary>
    void write_charts();

    /// <summary>
    ///
    /// </summary>
    void write_chartsheets();

    /// <summary>
    ///
    /// </summary>
    void write_worksheets();

    /// <summary>
    ///
    /// </summary>
    void write_external_links();

    /// <summary>
    /// A reference to the workbook which is the object of read/write operations.
    /// </summary>
    workbook &workbook_;

    /// <summary>
    /// The internal archive which holds files representing workbook_ during
    /// read/write operations.
    /// </summary>
    zip_file archive_;
};

} // namespace xlnt
