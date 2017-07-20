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

#include <iostream>
#include <memory>
#include <vector>

#pragma warning(push)
#pragma warning(disable: 4458)
#include <arrow/api.h>
#include <arrow/python/pyarrow.h>
#pragma warning(pop)

#include <Python.h> // must be included after Arrow

#include <detail/default_case.hpp>
#include <detail/unicode.hpp>
#include <python_streambuf.hpp>
#include <xlnt/cell/cell.hpp>
#include <xlnt/cell/cell_reference.hpp>
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

    default_case(std::unique_ptr<arrow::ArrayBuilder>(nullptrptr));
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

    default_case(arrow::Field("", arrow::nullptr()));
}

} // namespace xlnt

bool import_pyarrow()
{
    static bool imported = false;

    if (!imported)
    {
        if (arrow::py::import_pyarrow() != 0)
        {
            if (PyErr_Occurred() != nullptr)
            {
                PyErr_Print();
                PyErr_Clear();
            }
        }
        else
        {
            imported = true;
        }
    }

    return imported;
}

extern "C" {

PyObject *xlntpyarrow_xlsx2arrow(PyObject *self, PyObject *args, PyObject *kwargs)
{
    static const char *keywords[] = { "io", "sheetname", "header", "skiprows",
        "skip_footer", "index_col", "names", "converters", "dtype", "true_values",
        "false_values", "parse_cols", "squeeze", "na_values", "thousands",
        "keep_default_na", "verbose", "convert_float", nullptr };
    static auto keywords_nc = const_cast<char **>(keywords);

    PyObject *io = nullptr;
    PyObject *sheetname = nullptr;
    PyObject *header = nullptr;
    PyObject *skiprows = nullptr;
    auto skip_footer = 0;
    PyObject *index_col = nullptr;
    PyObject *names = nullptr;
    PyObject *converters = nullptr;
    PyObject *dtype = nullptr;
    PyObject *true_values = nullptr;
    PyObject *false_values = nullptr;
    PyObject *parse_cols = nullptr;
    auto squeeze = false;
    PyObject *na_values = nullptr;
    const char *thousands = nullptr;
    auto keep_default_va = false;
    auto verbose = false;
    auto convert_float = false;

    std::cout << "here" << std::endl;

    if (!PyArg_ParseTupleAndKeywords(args, kwargs, "O|OOOiOOOOOOOpOzppp", keywords_nc, 
        &io, &sheetname, &header, &skiprows, &skip_footer, &index_col, &names, 
        &converters, &dtype, &true_values, &false_values, &parse_cols, &squeeze,
        &na_values, &thousands, &keep_default_va, &verbose, &convert_float))
    {
        PyErr_Print();
        PyErr_Clear();
        Py_RETURN_NONE;
    }

    std::cout << "here2" << std::endl;

    if (!import_pyarrow())
    {
        Py_RETURN_NONE;
    }

    std::cout << "here3" << std::endl;

    // arg #1, io
    xlnt::python_streambuf file_buffer(io);
    std::istream file_stream(&file_buffer);

    xlnt::streaming_workbook_reader reader;
    reader.open(file_stream);

    std::cout << "here4" << std::endl;

    // arg #2, sheetname
    auto sheet_titles = reader.sheet_titles();
    auto sheet_title = sheet_titles.front();

    std::cout << "here5 " << sheet_title << std::endl;

    if (sheetname != nullptr)
    {
        std::cout << "sheetname" << std::endl;

        if (PyLong_Check(sheetname))
        {
            std::cout << "is long" << std::endl;
            // handle int sheetname
            auto sheet_index = PyLong_AsLong(sheetname);
            sheet_title = sheet_titles.at(sheet_index);
        }
        else if (PyUnicode_Check(sheetname))
        {
            std::cout << "is string" << std::endl;
            // handle string sheetname
            sheet_title = std::string(reinterpret_cast<char *>(PyUnicode_1BYTE_DATA(sheetname)));
        }
    }

    std::cout << sheet_title << std::endl;
    reader.begin_worksheet(sheet_title);

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
            continue;
        }
        else if (cell.row() == 2)
        {
            auto column_name = column_names.at(cell.column().index - 1);
            auto field = make_type_field(column_name, cell.data_type());
            fields.push_back(std::make_shared<arrow::Field>(field));
            columns.push_back(make_array_builder(cell.data_type()));
        }

        auto builder = columns.at(cell.column().index - 1).get();

        switch (cell.data_type())
        {
        case xlnt::cell::type::number:
        {
            auto typed_builder = static_cast<arrow::DoubleBuilder*>(builder);
            typed_builder->Append(0);
            break;
        }
        case xlnt::cell::type::inline_string:
        case xlnt::cell::type::shared_string:
        case xlnt::cell::type::error:
        case xlnt::cell::type::formula_string:
        case xlnt::cell::type::empty:
        {
            auto typed_builder = static_cast<arrow::StringBuilder*>(builder);
            typed_builder->Append(cell.value<std::string>());
            break;
        }
        case xlnt::cell::type::boolean:
        {
            auto typed_builder = static_cast<arrow::BooleanBuilder*>(builder);
            typed_builder->Append(cell.value<bool>());
            break;
        }
        case xlnt::cell::type::date:
        {
            auto typed_builder = static_cast<arrow::Date32Builder*>(builder);
            typed_builder->Append(cell.value<int>());
            break;
        }
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

    return arrow::py::wrap_table(table);
}

PyObject *xlntpyarrow_arrow2xlsx(PyObject *self, PyObject *args, PyObject *kwargs)
{
    static const char *keywords[] = { "table", "file", nullptr };
    static auto keywords_nc = const_cast<char **>(keywords);

    PyObject *table = nullptr;
    PyObject *file = nullptr;

    if (!PyArg_ParseTupleAndKeywords(args, kwargs, "OO", keywords_nc, &table, &file))
    {
        return nullptr;
    }

    if (!import_pyarrow())
    {
        Py_RETURN_NONE;
    }

    /*
    auto table = arrow::py::unwrap_table(pytable);
    xlnt::python_streambuf buffer(pyfile);
    std::ostream stream(&buffer);

    xlnt::streaming_workbook_writer writer;
    writer.open(s);

    writer.add_worksheet("Sheet1");

    for (auto i = 0; i < table->num_columns(); ++i)
    {
    auto column_name = table->schema()->field(i)->name();
    writer.add_cell(xlnt::cell_reference(i + 1, 1)).value(column_name);
    }
    */

    Py_RETURN_NONE;
}

} // extern "C"
