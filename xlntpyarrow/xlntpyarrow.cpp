#include <iostream>
#include <memory>
#include <vector>
#include <xlnt/utils/xlntarrow.hpp>
#include <python_streambuf.hpp>
#include <Python.h>

PyObject *xlsx2arrow(PyObject *file)
{
    xlnt::arrow::streambuf buffer(file);
    std::istream stream(&buffer);
    std::shared_ptr<arrow::Schema> schema;
    std::vector<std::shared_ptr<arrow::Column>> columns;
    arrow::Table table(schema, columns);
    xlnt::arrow::xlsx2arrow(stream, table);

    Py_RETURN_NONE;
}

extern "C" {

/*
 * Implements XLSX->pyarrow table function.
 */
PyDoc_STRVAR(xlntpyarrow_xlsx2arrow_doc, "xlsx2arrow(in_file)\
\
Returns an arrow table representing the given XLSX file object.");

PyObject *xlntpyarrow_xlsx2arrow(PyObject *self, PyObject *args, PyObject *kwargs)
{
    PyObject *file = nullptr;
    static const char *keywords[] = { "file", nullptr };
    static auto keywords_nc = const_cast<char **>(keywords);

    if (!PyArg_ParseTupleAndKeywords(args, kwargs, "O", keywords_nc, &file))
    {
        return nullptr;
    }

    return xlsx2arrow(file);
}


/*
* Implements pyarrow table->XLSX function.
*/
PyDoc_STRVAR(xlntpyarrow_arrow2xlsx_doc, "arrow2xlsx(table, out_file)\
\
Writes the given arrow table to out_file as an XLSX file.");

PyObject *xlntpyarrow_arrow2xlsx(PyObject *self, PyObject *args, PyObject *kwargs)
{
    PyObject *obj = nullptr;
    static const char *keywords[] = { "file", nullptr };
    static auto keywords_nc = const_cast<char **>(keywords);

    if (!PyArg_ParseTupleAndKeywords(args, kwargs, "Oi", keywords_nc, &obj))
    {
        return nullptr;
    }

    Py_RETURN_NONE;
}

static PyMethodDef xlntpyarrow_functions[] =
{
    { "xlsx2arrow", (PyCFunction)xlntpyarrow_xlsx2arrow, METH_VARARGS | METH_KEYWORDS, xlntpyarrow_xlsx2arrow_doc },
    { "arrow2xlsx", (PyCFunction)xlntpyarrow_arrow2xlsx, METH_VARARGS | METH_KEYWORDS, xlntpyarrow_arrow2xlsx_doc },
    { nullptr, nullptr, 0, nullptr }
};

int exec_xlntpyarrow(PyObject *module)
{
    PyModule_AddFunctions(module, xlntpyarrow_functions);

    PyModule_AddStringConstant(module, "__author__", "Thomas Fussell");
    PyModule_AddStringConstant(module, "__version__", "0.9.0");
    PyModule_AddIntConstant(module, "year", 2017);

    return 0;
}

PyDoc_STRVAR(xlntpyarrow_doc, "The xlntpyarrow module");

static PyModuleDef_Slot xlntpyarrow_slots[] =
{
    { Py_mod_exec, (void *)exec_xlntpyarrow },
    { 0, nullptr }
};

static PyModuleDef xlntpyarrow_def =
{
    PyModuleDef_HEAD_INIT,
    "xlntpyarrow",
    xlntpyarrow_doc,
    0,              /* m_size */
    nullptr,           /* m_methods */
    xlntpyarrow_slots,
    nullptr,           /* m_traverse */
    nullptr,           /* m_clear */
    nullptr,           /* m_free */
};

PyMODINIT_FUNC PyInit_xlntpyarrow()
{
    return PyModuleDef_Init(&xlntpyarrow_def);
}

} // extern "C"
