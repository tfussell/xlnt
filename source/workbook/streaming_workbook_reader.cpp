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

#include <detail/implementations/workbook_impl.hpp>
#include <detail/serialization/open_stream.hpp>
#include <detail/serialization/vector_streambuf.hpp>
#include <detail/serialization/xlsx_consumer.hpp>
#include <xlnt/cell/cell.hpp>
#include <xlnt/packaging/manifest.hpp>
#include <xlnt/utils/optional.hpp>
#include <xlnt/workbook/streaming_workbook_reader.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/worksheet/worksheet.hpp>

namespace xlnt {

streaming_workbook_reader::streaming_workbook_reader()
{

}

streaming_workbook_reader::~streaming_workbook_reader()
{
    close();
}

void streaming_workbook_reader::close()
{
    if (consumer_)
    {
        consumer_.reset(nullptr);
        stream_buffer_.reset(nullptr);
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

bool streaming_workbook_reader::has_worksheet(const std::string &name)
{
    auto titles = sheet_titles();
    return std::find(titles.begin(), titles.end(), name) != titles.end();
}

void streaming_workbook_reader::begin_worksheet(const std::string &title)
{
    if (!has_worksheet(title))
    {
        throw xlnt::exception("sheet not found");
    }

    worksheet_rel_id_ = workbook_->impl().sheet_title_rel_id_map_.at(title);
    const auto workbook_rel = workbook_->manifest()
        .relationship(path("/"), relationship_type::office_document);
    const auto worksheet_rel = workbook_->manifest()
        .relationship(workbook_rel.target().path(), worksheet_rel_id_);

    auto rel_chain = std::vector<relationship>{ workbook_rel, worksheet_rel };

    const auto &manifest = consumer_->target_.manifest();
    const auto part_path = manifest.canonicalize(rel_chain);
    auto part_stream_buffer = consumer_->archive_->open(part_path);
    part_stream_buffer_.swap(part_stream_buffer);
    part_stream_.reset(new std::istream(part_stream_buffer_.get()));
    parser_.reset(new xml::parser(*part_stream_, part_path.string()));
    consumer_->parser_ = parser_.get();

    consumer_->current_worksheet_ = nullptr;

    for (auto &impl : workbook_->impl().worksheets_)
    {
        if (impl.title_ == title)
        {
            consumer_->current_worksheet_ = &impl;
        }
    }

    if (consumer_->current_worksheet_ == nullptr)
    {
        throw xlnt::exception("sheet not found");
    }

    consumer_->read_worksheet_begin(worksheet_rel_id_);
}

worksheet streaming_workbook_reader::end_worksheet()
{
    return consumer_->read_worksheet_end(worksheet_rel_id_);
}

void streaming_workbook_reader::open(const std::vector<std::uint8_t> &data)
{
    stream_buffer_.reset(new detail::vector_istreambuf(data));
    stream_.reset(new std::istream(stream_buffer_.get()));
    open(*stream_);
}

void streaming_workbook_reader::open(const std::string &filename)
{
    stream_.reset(new std::ifstream());
    xlnt::detail::open_stream(static_cast<std::ifstream &>(*stream_), filename);
    open(*stream_);
}

#ifdef _MSC_VER
void streaming_workbook_reader::open(const std::wstring &filename)
{
    stream_.reset(new std::ifstream());
    xlnt::detail::open_stream(static_cast<std::ifstream &>(*stream_), filename);
    open(*stream_);
}
#endif

void streaming_workbook_reader::open(const xlnt::path &filename)
{
    stream_.reset(new std::ifstream());
    xlnt::detail::open_stream(static_cast<std::ifstream &>(*stream_), filename.string());
    open(*stream_);
}

void streaming_workbook_reader::open(std::istream &stream)
{
    workbook_.reset(new workbook());
    consumer_.reset(new detail::xlsx_consumer(*workbook_));
    consumer_->open(stream);

    const auto workbook_rel = workbook_->manifest()
        .relationship(path("/"), relationship_type::office_document);
    const auto workbook_path = workbook_rel.target().path();
}

void streaming_workbook_reader::open(std::unique_ptr<std::streambuf> &&buffer)
{
    stream_buffer_.swap(buffer);
    stream_.reset(new std::istream(stream_buffer_.get()));
    open(*stream_);
}

std::vector<std::string> streaming_workbook_reader::sheet_titles()
{
    return workbook_->sheet_titles();
}

} // namespace xlnt
