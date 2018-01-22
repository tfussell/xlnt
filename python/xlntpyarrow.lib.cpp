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

#include <exception>
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
            throw xlnt::exception("Import of pyarrow failed.");
        }

        imported = true;
    }
}

arrow::ArrayBuilder *make_array_builder(arrow::Type::type type)
{
    auto pool = arrow::default_memory_pool();
    auto builder = static_cast<arrow::ArrayBuilder *>(nullptr);

    switch(type)
    {
    case arrow::Type::NA:
        break;

    case arrow::Type::UINT8:
        builder = new arrow::TypeTraits<arrow::UInt8Type>::BuilderType(pool);
        break;

    case arrow::Type::INT8:
        builder = new arrow::TypeTraits<arrow::Int8Type>::BuilderType(pool);
        break;

    case arrow::Type::UINT16:
        builder = new arrow::TypeTraits<arrow::UInt16Type>::BuilderType(pool);
        break;

    case arrow::Type::INT16:
        builder = new arrow::TypeTraits<arrow::Int16Type>::BuilderType(pool);
        break;

    case arrow::Type::UINT32:
        builder = new arrow::TypeTraits<arrow::UInt32Type>::BuilderType(pool);
        break;

    case arrow::Type::INT32:
        builder = new arrow::TypeTraits<arrow::Int32Type>::BuilderType(pool);
        break;

    case arrow::Type::UINT64:
        builder = new arrow::TypeTraits<arrow::UInt64Type>::BuilderType(pool);
        break;

    case arrow::Type::INT64:
        builder = new arrow::TypeTraits<arrow::Int64Type>::BuilderType(pool);
        break;

    case arrow::Type::DATE64:
        builder = new arrow::TypeTraits<arrow::Date64Type>::BuilderType(pool);
        break;

    case arrow::Type::DATE32:
        builder = new arrow::TypeTraits<arrow::Date32Type>::BuilderType(pool);
        break;
/*
    case arrow::Type::TIMESTAMP:
        builder = new arrow::TypeTraits<arrow::TimestampType>::BuilderType(pool);
        break;

    case arrow::Type::TIME32:
        builder = new arrow::TypeTraits<arrow::Time32Type>::BuilderType(pool);
        break;

    case arrow::Type::TIME64:
        builder = new arrow::TypeTraits<arrow::Time64Type>::BuilderType(pool);
        break;
*/
    case arrow::Type::HALF_FLOAT:
        builder = new arrow::TypeTraits<arrow::HalfFloatType>::BuilderType(pool);
        break;

    case arrow::Type::FLOAT:
        builder = new arrow::TypeTraits<arrow::FloatType>::BuilderType(pool);
        break;

    case arrow::Type::DOUBLE:
        builder = new arrow::TypeTraits<arrow::DoubleType>::BuilderType(pool);
        break;
/*
    case arrow::Type::DECIMAL:
        builder = new arrow::TypeTraits<arrow::DecimalType>::BuilderType(pool, type);
        break;
*/
    case arrow::Type::BOOL:
        builder = new arrow::TypeTraits<arrow::BooleanType>::BuilderType(pool);
        break;

    case arrow::Type::STRING:
        builder = new arrow::TypeTraits<arrow::StringType>::BuilderType(pool);
        break;

    case arrow::Type::BINARY:
        builder = new arrow::TypeTraits<arrow::BinaryType>::BuilderType(pool);
        break;
/*
    case arrow::Type::FIXED_SIZE_BINARY:
        builder = new arrow::TypeTraits<arrow::FixedSizeBinaryType>::BuilderType(pool);
        break;

    case arrow::Type::LIST:
        builder = new arrow::TypeTraits<arrow::ListType>::BuilderType(pool);
        break;

    case arrow::Type::STRUCT:
        builder = new arrow::TypeTraits<arrow::StructType>::BuilderType(pool);
        break;

    case arrow::Type::UNION:
        builder = new arrow::TypeTraits<arrow::UnionType>::BuilderType(pool);
        break;

    case arrow::Type::DICTIONARY:
        builder = new arrow::TypeTraits<arrow::DictionaryType>::BuilderType(pool);
        break;
*/
    default:
        throw xlnt::exception("not implemented");
    }

    return builder;
}

void open_file(xlnt::streaming_workbook_reader &reader, pybind11::object file)
{
    reader.open(std::unique_ptr<std::streambuf>(new xlnt::python_streambuf(file)));
}

template<typename T>
T cell_value(xlnt::cell cell)
{
    return static_cast<T>(cell.value<double>());
}

// from https://stackoverflow.com/questions/1659440/32-bit-to-16-bit-floating-point-conversion
std::uint16_t float_to_half(float f)
{
    auto x = static_cast<std::uint32_t>(f);
    auto half = ((x >> 16) & 0x8000)
        | ((((x & 0x7f800000) - 0x38000000) >> 13) & 0x7c00)
        | ((x >> 13) & 0x03ff);

    return half;
}

void append_cell_value(arrow::ArrayBuilder *builder, arrow::Type::type type, xlnt::cell cell)
{
    const status = arrow::Status::OK();

    switch (type)
    {
    case arrow::Type::NA:
        break;

    case arrow::Type::BOOL:
        status = static_cast<arrow::BooleanBuilder *>(builder)
            ->Append(cell.value<bool>());
        break;

    case arrow::Type::UINT8:
        status = static_cast<arrow::UInt8Builder *>(builder)
            ->Append(cell_value<std::uint8_t>(cell));
        break;

    case arrow::Type::INT8:
        status = static_cast<arrow::Int8Builder *>(builder)
          ->Append(cell_value<std::uint8_t>(cell));
        break;

    case arrow::Type::UINT16:
        status = static_cast<arrow::UInt16Builder *>(builder)
            ->Append(cell_value<std::uint16_t>(cell));
        break;

    case arrow::Type::INT16:
        status = static_cast<arrow::Int16Builder *>(builder)
            ->Append(cell_value<std::int16_t>(cell));
        break;

    case arrow::Type::UINT32:
        status = static_cast<arrow::UInt32Builder *>(builder)
            ->Append(cell_value<std::uint32_t>(cell));
        break;

    case arrow::Type::INT32:
        status = static_cast<arrow::Int32Builder *>(builder)
            ->Append(cell_value<std::int32_t>(cell));
        break;

    case arrow::Type::UINT64:
        status = static_cast<arrow::UInt64Builder *>(builder)
            ->Append(cell_value<std::uint64_t>(cell));
        break;

    case arrow::Type::INT64:
        status = static_cast<arrow::Int64Builder *>(builder)
            ->Append(cell_value<std::int64_t>(cell));
        break;

    case arrow::Type::HALF_FLOAT:
        status = static_cast<arrow::HalfFloatBuilder *>(builder)
            ->Append(float_to_half(cell_value<float>(cell)));
        break;

    case arrow::Type::FLOAT:
        status = static_cast<arrow::FloatBuilder *>(builder)
            ->Append(cell_value<float>(cell));
        break;

    case arrow::Type::DOUBLE:
        status = static_cast<arrow::DoubleBuilder *>(builder)
            ->Append(cell_value<double>(cell));
        break;

    case arrow::Type::STRING:
        status = static_cast<arrow::StringBuilder *>(builder)
            ->Append(cell.value<std::string>());
        break;

    case arrow::Type::BINARY:
        status = static_cast<arrow::BinaryBuilder *>(builder)
            ->Append(cell.value<std::string>());
        break;

    case arrow::Type::FIXED_SIZE_BINARY:
        status = static_cast<arrow::FixedSizeBinaryBuilder *>(builder)
            ->Append(cell.value<std::string>());
        break;

    case arrow::Type::DATE32:
        status = static_cast<arrow::Date32Builder *>(builder)
            ->Append(cell_value<arrow::Date32Type::c_type>(cell));
        break;

    case arrow::Type::DATE64:
        status = static_cast<arrow::Date64Builder *>(builder)
            ->Append(cell_value<arrow::Date64Type::c_type>(cell));
        break;

    case arrow::Type::TIMESTAMP:
        status = static_cast<arrow::TimestampBuilder *>(builder)
            ->Append(cell_value<arrow::TimestampType::c_type>(cell));
        break;

    case arrow::Type::TIME32:
        status = static_cast<arrow::Time32Builder *>(builder)
            ->Append(cell_value<arrow::Time32Type::c_type>(cell));
        break;

    case arrow::Type::TIME64:
        status = static_cast<arrow::Time64Builder *>(builder)
            ->Append(cell_value<arrow::Time64Type::c_type>(cell));
        break;
/*
    case arrow::Type::INTERVAL:
        status = static_cast<arrow::IntervalBuilder *>(builder)
            ->Append(cell_value<std::int64_t>(cell));
        break;

    case arrow::Type::DECIMAL:
        status = static_cast<arrow::DecimalBuilder *>(builder)
            ->Append(cell.value<std::string>());
        break;

    case arrow::Type::LIST:
        status = static_cast<arrow::ListBuilder *>(builder)
            ->Append(cell.value<std::string>());
        break;

    case arrow::Type::STRUCT:
        status = static_cast<arrow::StructBuilder *>(builder)
            ->Append(cell.value<std::string>());
        break;

    case arrow::Type::UNION:
        status = static_cast<arrow::UnionBuilder *>(builder)
            ->Append(cell.value<std::string>());
        break;

    case arrow::Type::DICTIONARY:
        status = static_cast<arrow::DictionaryBuilder *>(builder)
            ->Append(cell.value<std::string>());
        break;
*/
    default:
        throw xlnt::exception("not implemented");
    }

    if (status != arrow::Status::OK())
    {
        throw xlnt::exception("Append failed");
    }
}

pybind11::handle read_batch(xlnt::streaming_workbook_reader &reader,
    pybind11::object pyschema, int max_rows)
{
    import_pyarrow();

    std::shared_ptr<arrow::Schema> schema;
    arrow::py::unwrap_schema(pyschema.ptr(), &schema);

    std::vector<arrow::Type::type> column_types;

    for (auto i = 0; i < schema->num_fields(); ++i)
    {
        column_types.push_back(schema->field(i)->type()->id());
    }

    auto builders = std::vector<std::unique_ptr<arrow::ArrayBuilder>>();

    for (auto type : column_types)
    {
        builders.emplace_back(make_array_builder(type));
    }

    auto row = std::int64_t(0);

    while (row < max_rows)
    {
        if (!reader.has_cell()) break;

        for (auto column = 0; column < schema->num_fields(); ++column)
        {
            if (!reader.has_cell()) break;

            auto cell = reader.read_cell();
            auto zero_indexed_column = cell.column().index - 1;
            auto column_type = column_types.at(zero_indexed_column);
            auto builder = builders.at(zero_indexed_column).get();

            append_cell_value(builder, column_type, cell);
        }

        ++row;
    }

    auto columns = std::vector<std::shared_ptr<arrow::Array>>();

    for (auto &builder : builders)
    {
        std::shared_ptr<arrow::Array> column;
        builder->Finish(&column);
        columns.emplace_back(column);
    }

    auto batch_pointer = std::make_shared<arrow::RecordBatch>(schema, row, columns);
    auto batch_object = arrow::py::wrap_record_batch(batch_pointer);
    auto batch_handle = pybind11::handle(batch_object); // don't need to incr. reference count, right?

    return batch_handle;
}

PYBIND11_MODULE(lib, m)
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
        .def("value_bool", [](xlnt::cell &cell)
        {
            return cell.value<bool>();
        })
        .def("value_unsigned_int", [](xlnt::cell &cell)
        {
            return cell.value<unsigned int>();
        })
        .def("value_double", [](xlnt::cell &cell)
        {
            return cell.value<double>();
        })
        .def("data_type", [](xlnt::cell &cell)
            {
                return cell.data_type();
            })
        .def("row", &xlnt::cell::row)
        .def("column", [](xlnt::cell &cell)
            {
                return cell.column().index;
            })
        .def("format_is_date", [](xlnt::cell &cell)
            {
                return cell.has_format() && cell.number_format().is_date_format();
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
