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

#include <detail/serialization/xlsx_consumer.hpp>
#include <xlnt/workbook/streaming_workbook_reader.hpp>

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

void streaming_workbook_reader::on_cell(std::function<void(cell)> callback)
{
    //consumer_->on_cell(callback);
}

void streaming_workbook_reader::on_worksheet_start(std::function<void(std::string)> callback)
{
    //consumer_->on_worksheet_start(callback);
}

void streaming_workbook_reader::on_worksheet_end(std::function<void(worksheet)> callback)
{
    //consumer_->on_worksheet_end(callback);
}

void streaming_workbook_reader::open(const std::vector<std::uint8_t> &data)
{
 
}

void streaming_workbook_reader::open(const std::string &filename)
{

}

#ifdef _MSC_VER
void streaming_workbook_reader::open(const std::wstring &filename)
{

}
#endif

void streaming_workbook_reader::open(const xlnt::path &filename)
{

}

void streaming_workbook_reader::open(std::istream &stream)
{

}

} // namespace xlnt
