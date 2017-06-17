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

#include <fstream>

#include <detail/serialization/vector_streambuf.hpp>
#include <detail/serialization/xlsx_consumer.hpp>
#include <xlnt/cell/cell.hpp>
#include <xlnt/utils/optional.hpp>
#include <xlnt/workbook/streaming_workbook_reader.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/worksheet/worksheet.hpp>


namespace {

//TODO: (important) this is duplicated from workbook.cpp, find a common place to keep it
#ifdef _MSC_VER
void open_stream(std::ifstream &stream, const std::wstring &path)
{
    stream.open(path, std::ios::binary);
}

void open_stream(std::ofstream &stream, const std::wstring &path)
{
    stream.open(path, std::ios::binary);
}

void open_stream(std::ifstream &stream, const std::string &path)
{
    open_stream(stream, xlnt::path(path).wstring());
}

void open_stream(std::ofstream &stream, const std::string &path)
{
    open_stream(stream, xlnt::path(path).wstring());
}
#else
void open_stream(std::ifstream &stream, const std::string &path)
{
    stream.open(path, std::ios::binary);
}

void open_stream(std::ofstream &stream, const std::string &path)
{
    stream.open(path, std::ios::binary);
}
#endif

} // namespace


namespace xlnt {

streaming_workbook_reader::~streaming_workbook_reader()
{
  close();
}

void streaming_workbook_reader::close()
{
    if (consumer_)
    {
        consumer_.reset(nullptr);
    }
}

bool streaming_workbook_reader::has_cell()
{
    return consumer_->has_cell();
}

cell streaming_workbook_reader::read_cell()
{
    return consumer_->read_cell();
}

bool streaming_workbook_reader::has_worksheet()
{
    return consumer_->has_worksheet();
}

std::string streaming_workbook_reader::begin_worksheet()
{
    return consumer_->read_worksheet_begin();
}

worksheet streaming_workbook_reader::end_worksheet()
{
    return consumer_->end_worksheet();
}

void streaming_workbook_reader::open(const std::vector<std::uint8_t> &data)
{
    detail::vector_istreambuf buffer(data);
    std::istream buffer_stream(&buffer);
    open(buffer_stream);
}

void streaming_workbook_reader::open(const std::string &filename)
{
    std::ifstream file_stream;
    open_stream(file_stream, filename);
}

#ifdef _MSC_VER
void streaming_workbook_reader::open(const std::wstring &filename)
{
    std::ifstream file_stream;
    open_stream(file_stream, filename);
}
#endif

void streaming_workbook_reader::open(const xlnt::path &filename)
{
    std::ifstream file_stream;
    open_stream(file_stream, filename.string());
}

void streaming_workbook_reader::open(std::istream &stream)
{
    workbook_.reset(new workbook());
    consumer_.reset(new detail::xlsx_consumer(*workbook_));
    consumer_->open(stream);
}

} // namespace xlnt
