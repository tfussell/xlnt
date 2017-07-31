// Copyright (c) 2014-2017 Thomas Fussell
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

#include <functional>
#include <iostream>
#include <memory>
#include <string>
#include <vector>

#include <xlnt/xlnt_config.hpp>

namespace xml {
class parser;
}

namespace xlnt {

class cell;
template<typename T>
class optional;
class path;
class workbook;
class worksheet;

namespace detail {
class xlsx_consumer;
}

/// <summary>
/// workbook is the container for all other parts of the document.
/// </summary>
class XLNT_API streaming_workbook_reader
{
public:
    streaming_workbook_reader();
    ~streaming_workbook_reader();

    /// <summary>
    /// Closes currently open read stream. This will be called automatically
    /// by the destructor if it hasn't already been called manually.
    /// </summary>
    void close();

    bool has_cell();

    /// <summary>
    /// Reads the next cell in the current worksheet and optionally returns it if
    /// the last cell in the sheet has not yet been read.
    /// </summary>
    cell read_cell();

    bool has_worksheet(const std::string &name);

    /// <summary>
    /// Beings reading of the next worksheet in the workbook and optionally
    /// returns its title if the last worksheet has not yet been read.
    /// </summary>
    void begin_worksheet(const std::string &name);

    /// <summary>
    /// Ends reading of the current worksheet in the workbook and optionally
    /// returns a worksheet object corresponding to the worksheet with the title
    /// returned by begin_worksheet().
    /// </summary>
    worksheet end_worksheet();

    /// <summary>
    /// Interprets byte vector data as an XLSX file and sets the content of this
    /// workbook to match that file.
    /// </summary>
    void open(const std::vector<std::uint8_t> &data);

    /// <summary>
    /// Interprets file with the given filename as an XLSX file and sets
    /// the content of this workbook to match that file.
    /// </summary>
    void open(const std::string &filename);

#ifdef _MSC_VER
    /// <summary>
    /// Interprets file with the given filename as an XLSX file and sets
    /// the content of this workbook to match that file.
    /// </summary>
    void open(const std::wstring &filename);
#endif

    /// <summary>
    /// Interprets file with the given filename as an XLSX file and sets the
    /// content of this workbook to match that file.
    /// </summary>
    void open(const path &filename);

    /// <summary>
    /// Interprets data in stream as an XLSX file and sets the content of this
    /// workbook to match that file.
    /// </summary>
    void open(std::istream &stream);

    /// <summary>
    /// Holds the given streambuf internally, creates a std::istream backed
    /// by the given buffer, and calls open(std::istream &) with that stream.
    /// </summary>
    void open(std::unique_ptr<std::streambuf> &&buffer);

    /// <summary>
    /// Returns a vector of the titles of sheets in the workbook in order.
    /// </summary>
    std::vector<std::string> sheet_titles();

private:
    std::string worksheet_rel_id_;
    std::unique_ptr<detail::xlsx_consumer> consumer_;
    std::unique_ptr<workbook> workbook_;
    std::unique_ptr<std::istream> stream_;
    std::unique_ptr<std::streambuf> stream_buffer_;
    std::unique_ptr<std::istream> part_stream_;
    std::unique_ptr<std::streambuf> part_stream_buffer_;
    std::unique_ptr<xml::parser> parser_;
};

} // namespace xlnt
