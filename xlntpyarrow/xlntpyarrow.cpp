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

#include <arrow/api.h>
#include <arrow/python/pyarrow.h>
#include <pybind11/pybind11.h>
#include <pybind11/stl.h>
#include <xlnt/xlnt.hpp>
#include <xlnt/workbook/streaming_workbook_reader.hpp>
#include <python_streambuf.hpp>

void import_pyarrow()
{
    static auto imported = false;

    if (!imported)
    {
        if (arrow::py::import_pyarrow() != 0)
        {
            throw std::exception("Import of pyarrow failed.");
        }

        imported = true;
    }
}

std::vector<arrow::Type::type> extract_schema_types(std::shared_ptr<arrow::Schema> &schema)
{
    auto types = std::vector<arrow::Type::type>();

    for (auto i = 0; i < schema->num_fields(); ++i)
    {
        types.push_back(schema->field(i)->type()->id());
    }

    return types;
}

std::unique_ptr<arrow::ArrayBuilder> make_array_builder(arrow::Type::type type)
{
    std::unique_ptr<arrow::ArrayBuilder> builder;
    auto pool = arrow::default_memory_pool();

    switch (type)
    {
    case arrow::Type::NA:
        break;

    case arrow::Type::BOOL:
        builder.reset(new arrow::BooleanBuilder(pool));
        break;

    case arrow::Type::UINT8:
        break;

    case arrow::Type::INT8:
        break;

    case arrow::Type::UINT16:
        break;

    case arrow::Type::INT16:
        break;

    case arrow::Type::UINT32:
        break;

    case arrow::Type::INT32:
        break;

    case arrow::Type::UINT64:
        break;

    case arrow::Type::INT64:
        break;

    case arrow::Type::HALF_FLOAT:
        break;

    case arrow::Type::FLOAT:
        break;

    case arrow::Type::DOUBLE:
        builder.reset(new arrow::DoubleBuilder(pool));
        break;

    case arrow::Type::STRING:
        builder.reset(new arrow::StringBuilder(pool));
        break;

    case arrow::Type::BINARY:
        break;

    case arrow::Type::FIXED_SIZE_BINARY:
        break;

    case arrow::Type::DATE32:
        builder.reset(new arrow::Date32Builder(pool));
        break;

    case arrow::Type::DATE64:
        break;

    case arrow::Type::TIMESTAMP:
        break;

    case arrow::Type::TIME32:
        break;

    case arrow::Type::TIME64:
        break;

    case arrow::Type::INTERVAL:
        break;

    case arrow::Type::DECIMAL:
        break;

    case arrow::Type::LIST:
        break;

    case arrow::Type::STRUCT:
        break;

    case arrow::Type::UNION:
        break;

    case arrow::Type::DICTIONARY:
        break;
    }

    return builder;
}

void open_file(xlnt::streaming_workbook_reader &reader, pybind11::object file)
{
    reader.open(std::unique_ptr<std::streambuf>(new xlnt::python_streambuf(file)));
}

pybind11::handle read_batch(xlnt::streaming_workbook_reader &reader,
    pybind11::object pyschema, int max_rows)
{
    import_pyarrow();

    std::shared_ptr<arrow::Schema> schema;
    arrow::py::unwrap_schema(pyschema.ptr(), &schema);

    auto column_types = extract_schema_types(schema);
    auto builders = std::vector<std::shared_ptr<arrow::ArrayBuilder>>();
    auto num_rows = std::int64_t(0);

    for (auto type : column_types)
    {
        builders.push_back(make_array_builder(type));
    }

    for (auto row = 0; row < max_rows; ++row)
    {
        if (!reader.has_cell()) break;

        if (row % 1000 == 0)
        {
            std::cout << row << std::endl;
        }

        for (auto column = 0; column < schema->num_fields(); ++column)
        {
            if (!reader.has_cell()) break;

            auto cell = reader.read_cell();
            auto column_type = column_types.at(column);
            auto builder = builders.at(cell.column().index - 1).get();

            switch (column_type)
            {
            case arrow::Type::NA:
                break;

            case arrow::Type::BOOL:
                static_cast<arrow::BooleanBuilder *>(builder)->Append(cell.value<bool>());
                break;

            case arrow::Type::UINT8:
                break;

            case arrow::Type::INT8:
                break;

            case arrow::Type::UINT16:
                break;

            case arrow::Type::INT16:
                break;

            case arrow::Type::UINT32:
                break;

            case arrow::Type::INT32:
                break;

            case arrow::Type::UINT64:
                break;

            case arrow::Type::INT64:
                break;

            case arrow::Type::HALF_FLOAT:
                break;

            case arrow::Type::FLOAT:
                break;

            case arrow::Type::DOUBLE:
                static_cast<arrow::DoubleBuilder *>(builder)->Append(cell.value<long double>());
                break;

            case arrow::Type::STRING:
                static_cast<arrow::StringBuilder *>(builder)->Append(cell.value<std::string>());
                break;

            case arrow::Type::BINARY:
                break;

            case arrow::Type::FIXED_SIZE_BINARY:
                break;

            case arrow::Type::DATE32:
                static_cast<arrow::Date32Builder *>(builder)->Append(cell.value<int>());
                break;

            case arrow::Type::DATE64:
                break;

            case arrow::Type::TIMESTAMP:
                break;

            case arrow::Type::TIME32:
                break;

            case arrow::Type::TIME64:
                break;

            case arrow::Type::INTERVAL:
                break;

            case arrow::Type::DECIMAL:
                break;

            case arrow::Type::LIST:
                break;

            case arrow::Type::STRUCT:
                break;

            case arrow::Type::UNION:
                break;

            case arrow::Type::DICTIONARY:
                break;
            }
        }

        ++num_rows;
    }

    auto columns = std::vector<std::shared_ptr<arrow::Array>>();

    for (auto &builder : builders)
    {
        std::shared_ptr<arrow::Array> column;
        builder->Finish(&column);
        columns.emplace_back(column);
    }

    auto batch_pointer = std::make_shared<arrow::RecordBatch>(schema, num_rows, columns);
    auto batch_object = arrow::py::wrap_record_batch(batch_pointer);
    auto batch_handle = pybind11::handle(batch_object); // don't need to incr. reference count, right?

    return batch_handle;
}

PYBIND11_MODULE(xlntpyarrow, m)
{
    m.doc() = "streaming read/write interface for C++ XLSX library xlnt";

    pybind11::class_<xlnt::streaming_workbook_reader>(m, "StreamingWorkbookReader")
        .def(pybind11::init<>())
        .def("has_cell", &xlnt::streaming_workbook_reader::has_cell)
        .def("read_cell", &xlnt::streaming_workbook_reader::read_cell)
        .def("has_worksheet", &xlnt::streaming_workbook_reader::has_worksheet)
        .def("begin_worksheet", &xlnt::streaming_workbook_reader::begin_worksheet)
        .def("end_worksheet", &xlnt::streaming_workbook_reader::end_worksheet)
        .def("sheet_titles", &xlnt::streaming_workbook_reader::sheet_titles)
        .def("open", &open_file)
        .def("read_batch", &read_batch);

    pybind11::class_<xlnt::worksheet>(m, "Worksheet");

    pybind11::class_<xlnt::cell> cell(m, "Cell");
    cell.def("value_string", [](xlnt::cell &cell)
        {
            return cell.value<std::string>();
        })
        .def("data_type", [](xlnt::cell &cell)
            {
                return cell.data_type();
            })
        .def("row", &xlnt::cell::row)
        .def("column", [](xlnt::cell &cell)
            {
                return cell.column().index;
            });

    pybind11::enum_<xlnt::cell::type>(cell, "Type")
        .value("Empty", xlnt::cell::type::empty)
        .value("Boolean", xlnt::cell::type::boolean)
        .value("Date", xlnt::cell::type::date)
        .value("Error", xlnt::cell::type::error)
        .value("InlineString", xlnt::cell::type::inline_string)
        .value("Number", xlnt::cell::type::number)
        .value("SharedString", xlnt::cell::type::shared_string)
        .value("FormulaString", xlnt::cell::type::formula_string);
}
