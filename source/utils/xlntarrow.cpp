// Copyright (c) 2017 Thomas Fussell
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

#pragma warning(push)
#pragma warning(disable: 4458)
#include <arrow/api.h>
#pragma warning(pop)

#include <xlnt/cell/cell.hpp>
#include <xlnt/cell/cell_reference.hpp>
#include <xlnt/utils/xlntarrow.hpp>
#include <xlnt/workbook/streaming_workbook_reader.hpp>
#include <xlnt/workbook/streaming_workbook_writer.hpp>
#include <xlnt/worksheet/worksheet.hpp>

namespace {

std::unique_ptr<arrow::ArrayBuilder> make_array_builder(xlnt::cell::type type)
{
    switch (type) 
    {
    case xlnt::cell::type::number:
        return std::unique_ptr<arrow::ArrayBuilder>(new arrow::DoubleBuilder(arrow::default_memory_pool(), arrow::float64()));
    case xlnt::cell::type::inline_string:
    case xlnt::cell::type::shared_string:
    case xlnt::cell::type::error:
    case xlnt::cell::type::formula_string:
    case xlnt::cell::type::empty:
        return std::unique_ptr<arrow::StringBuilder>(new arrow::StringBuilder(arrow::default_memory_pool()));
    case xlnt::cell::type::boolean:
        return std::unique_ptr<arrow::ArrayBuilder>(new arrow::BooleanBuilder(arrow::default_memory_pool(), std::make_shared<arrow::BooleanType>()));
    case xlnt::cell::type::date:
        return std::unique_ptr<arrow::Date32Builder>(new arrow::Date32Builder(arrow::default_memory_pool()));
    }
}

arrow::Field make_type_field(const std::string &name, xlnt::cell::type type)
{
    switch (type)
    {
    case xlnt::cell::type::number:
        return arrow::Field(name, arrow::float64());
    case xlnt::cell::type::inline_string:
    case xlnt::cell::type::shared_string:
    case xlnt::cell::type::error:
    case xlnt::cell::type::formula_string:
    case xlnt::cell::type::empty:
        return arrow::Field(name, std::make_shared<arrow::StringType>());
    case xlnt::cell::type::boolean:
        return arrow::Field(name, arrow::boolean());
    case xlnt::cell::type::date:
        return arrow::Field(name, arrow::date32());
    }
}

} // namespace

namespace xlnt {

std::shared_ptr<arrow::Table> XLNT_API xlsx2arrow(std::istream &s)
{
    xlnt::streaming_workbook_reader reader;
    reader.open(s);

    reader.begin_worksheet();

    auto column_names = std::vector<std::string>();
    auto columns = std::vector<std::unique_ptr<arrow::ArrayBuilder>>();
    auto fields = std::vector<std::shared_ptr<arrow::Field>>();

    auto arrow_check = [](arrow::Status s)
    {
        if (!s.ok())
        {
            throw xlnt::exception("conversion error");
        }
    };

    while (reader.has_cell())
    {
        auto cell = reader.read_cell();

        if (cell.row() == 1)
        {
            column_names.push_back(cell.value<std::string>());
        }
        else if (cell.row() == 2)
        {
            auto column_name = column_names.at(cell.column().index - 1);
            auto field = make_type_field(column_name, cell.data_type());
            fields.push_back(std::make_shared<arrow::Field>(field));
            columns.push_back(make_array_builder(cell.data_type()));
        }
    }

    reader.end_worksheet();

    auto schema = std::make_shared<arrow::Schema>(fields);
    auto arrays = std::vector<std::shared_ptr<arrow::Array>>();

    for (size_t i = 0; i != columns.size(); ++i)
    {
        std::shared_ptr<arrow::Array> array;
        columns[i]->Finish(&array);
        arrays.emplace_back(array);
    }

    std::shared_ptr<arrow::Table> table;
    arrow_check(MakeTable(schema, arrays, &table));

    return table;
}

void XLNT_API arrow2xlsx(std::shared_ptr<const arrow::Table> &table, std::ostream &s)
{
  xlnt::streaming_workbook_writer writer;
  writer.open(s);

  writer.add_worksheet("Sheet1");

  for (auto i = 0; i < table->num_columns(); ++i)
  {
      auto column_name = table->schema()->field(i)->name();
      writer.add_cell(xlnt::cell_reference(i + 1, 1)).value(column_name);
  }
}

} // namespace xlnt
