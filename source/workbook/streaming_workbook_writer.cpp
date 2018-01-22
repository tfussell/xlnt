// Copyright (c) 2017-2018 Thomas Fussell
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

#include <detail/implementations/cell_impl.hpp>
#include <detail/implementations/worksheet_impl.hpp>
#include <detail/serialization/open_stream.hpp>
#include <detail/serialization/vector_streambuf.hpp>
#include <detail/serialization/xlsx_producer.hpp>
#include <xlnt/cell/cell.hpp>
#include <xlnt/packaging/manifest.hpp>
#include <xlnt/utils/optional.hpp>
#include <xlnt/workbook/streaming_workbook_writer.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/worksheet/worksheet.hpp>

namespace xlnt {

streaming_workbook_writer::streaming_workbook_writer()
{

}

streaming_workbook_writer::~streaming_workbook_writer()
{
    close();
}

void streaming_workbook_writer::close()
{
    if (producer_)
    {
        producer_.reset(nullptr);
        stream_buffer_.reset(nullptr);
    }
}

cell streaming_workbook_writer::add_cell(const cell_reference &ref)
{
    return producer_->add_cell(ref);
}

worksheet streaming_workbook_writer::add_worksheet(const std::string &title)
{
    return producer_->add_worksheet(title);
}

void streaming_workbook_writer::open(std::vector<std::uint8_t> &data)
{
    stream_buffer_.reset(new detail::vector_ostreambuf(data));
    stream_.reset(new std::ostream(stream_buffer_.get()));
    open(*stream_);
}

void streaming_workbook_writer::open(const std::string &filename)
{
    stream_.reset(new std::ofstream());
    xlnt::detail::open_stream(static_cast<std::ofstream &>(*stream_), filename);
    open(*stream_);
}

#ifdef _MSC_VER
void streaming_workbook_writer::open(const std::wstring &filename)
{
    stream_.reset(new std::ofstream());
    xlnt::detail::open_stream(static_cast<std::ofstream &>(*stream_), filename);
    open(*stream_);
}
#endif

void streaming_workbook_writer::open(const xlnt::path &filename)
{
    stream_.reset(new std::ofstream());
    xlnt::detail::open_stream(static_cast<std::ofstream &>(*stream_), filename.string());
    open(*stream_);
}

void streaming_workbook_writer::open(std::ostream &stream)
{
    workbook_.reset(new workbook());
    producer_.reset(new detail::xlsx_producer(*workbook_));
    producer_->open(stream);
    producer_->current_worksheet_ = new detail::worksheet_impl(workbook_.get(), 1, "Sheet1");
    producer_->current_cell_ = new detail::cell_impl();
    producer_->current_cell_->parent_ = producer_->current_worksheet_;
}

} // namespace xlnt
