// Copyright (c) 2014-2021 Thomas Fussell
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
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE
//
// @license: http://www.opensource.org/licenses/mit-license.php
// @author: see AUTHORS file
#pragma once

#include <cstddef>
#include <iterator>
#include <string>
#include <vector>

#include <xlnt/xlnt_config.hpp>

namespace xml {
class serializer;
}

namespace xlnt {

class cell;
class cell_reference;
class worksheet;

namespace detail {
class xlsx_producer;
} // namespace detail

/// <summary>
/// workbook is the container for all other parts of the document.
/// </summary>
class XLNT_API streaming_workbook_writer
{
public:
    streaming_workbook_writer();
    ~streaming_workbook_writer();

    /// <summary>
    /// Finishes writing of the remaining contents of the workbook and closes
    /// currently open write stream. This will be called automatically by the
    /// destructor if it hasn't already been called manually.
    /// </summary>
    void close();

    /// <summary>
    /// Writes a cell to the currently active worksheet at the position given by
    /// ref and with the given value. ref should be to the right of or below
    /// the previously written cell.
    /// </summary>
    cell add_cell(const cell_reference &ref);

    /// <summary>
    /// Ends writing of data to the current sheet and begins writing a new sheet
    /// with the given title.
    /// </summary>
    worksheet add_worksheet(const std::string &title);

    /// <summary>
    /// Serializes the workbook into an XLSX file and saves the bytes into
    /// byte vector data.
    /// </summary>
    void open(std::vector<std::uint8_t> &data);

    /// <summary>
    /// Serializes the workbook into an XLSX file and saves the data into a file
    /// named filename.
    /// </summary>
    void open(const std::string &filename);

#ifdef _MSC_VER
    /// <summary>
    /// Serializes the workbook into an XLSX file and saves the data into a file
    /// named filename.
    /// </summary>
    void open(const std::wstring &filename);
#endif

    /// <summary>
    /// Serializes the workbook into an XLSX file and saves the data into a file
    /// named filename.
    /// </summary>
    void open(const xlnt::path &filename);

    /// <summary>
    /// Serializes the workbook into an XLSX file and saves the data into stream.
    /// </summary>
    void open(std::ostream &stream);

    std::unique_ptr<xlnt::detail::xlsx_producer> producer_;
    std::unique_ptr<workbook> workbook_;
    std::unique_ptr<std::ostream> stream_;
    std::unique_ptr<std::streambuf> stream_buffer_;
    std::unique_ptr<std::ostream> part_stream_;
    std::unique_ptr<std::streambuf> part_stream_buffer_;
    std::unique_ptr<xml::serializer> serializer_;
};

} // namespace xlnt
