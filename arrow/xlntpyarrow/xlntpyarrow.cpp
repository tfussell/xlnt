#include <iostream>
#include <memory>
#include <vector>
#include <xlntarrow.hpp>
#include <python_streambuf.hpp>
#include <Python.h>

PyObject *xlsx2arrow(PyObject *file)
{
    boost::python::handle<> boost_file_handle(file);
    boost::python::object boost_file(boost_file_handle);
    boost_adaptbx::python::streambuf buffer(boost_file);
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
    PyObject *file = NULL;
    static char* keywords[] = { "file", NULL };

    if (!PyArg_ParseTupleAndKeywords(args, kwargs, "O", keywords, &file))
    {
        return NULL;
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
    PyObject *obj = NULL;
    static char* keywords[] = { "obj", NULL };

    if (!PyArg_ParseTupleAndKeywords(args, kwargs, "Oi", keywords, &obj))
    {
        return NULL;
    }

    Py_RETURN_NONE;
}

static PyMethodDef xlntpyarrow_functions[] =
{
    { "xlsx2arrow", (PyCFunction)xlntpyarrow_xlsx2arrow, METH_VARARGS | METH_KEYWORDS, xlntpyarrow_xlsx2arrow_doc },
    { "arrow2xlsx", (PyCFunction)xlntpyarrow_arrow2xlsx, METH_VARARGS | METH_KEYWORDS, xlntpyarrow_arrow2xlsx_doc },
    { NULL, NULL, 0, NULL } /* marks end of array */
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
    { 0, NULL }
};

static PyModuleDef xlntpyarrow_def =
{
    PyModuleDef_HEAD_INIT,
    "xlntpyarrow",
    xlntpyarrow_doc,
    0,              /* m_size */
    NULL,           /* m_methods */
    xlntpyarrow_slots,
    NULL,           /* m_traverse */
    NULL,           /* m_clear */
    NULL,           /* m_free */
};

PyMODINIT_FUNC PyInit_xlntpyarrow()
{
    return PyModuleDef_Init(&xlntpyarrow_def);
}

} // extern "C"
