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
#include <vector>

#include <xlnt/xlnt_config.hpp>

namespace xlnt {

class cell;
class path;
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
    ~streaming_workbook_reader();

    /// <summary>
    /// Closes currently open read stream. This will be called automatically
    /// by the destructor if it hasn't already been called manually.
    /// </summary>
    void close();

    /// <summary>
    /// Registers callback as the function to be called when a cell is read.
    /// </summary>
    void on_cell(std::function<void(cell)> callback);

    /// <summary>
    /// Registers callback as the function to be called when a worksheet is
    /// opened for reading. The callback will be called with the title of the
    /// worksheet.
    /// </summary>
    void on_worksheet_start(std::function<void(std::string)> callback);

    /// <summary>
    /// Registers callback as the function to be called when a worksheet is
    /// finished being read. The callback will be called with the worksheet
    /// object (which will not contain any cells). Cells should be handled
    /// via registering a callback with on_cell.
    /// </summary>
    void on_worksheet_end(std::function<void(worksheet)> callback);

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

private:
    std::unique_ptr<xlnt::detail::xlsx_consumer> consumer_;
};

} // namespace xlnt
