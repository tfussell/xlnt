#include <iostream>
#include <memory>
#include <vector>

#include <arrow/api.h>
#include <arrow/python/pyarrow.h>
#include <Python.h> // must be included after Arrow

#include <python_streambuf.hpp>
#include <xlnt/utils/xlntarrow.hpp>

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

PyObject *xlsx2arrow(PyObject *pyfile)
{
    if (!import_pyarrow())
    {
        Py_RETURN_NONE;
    }

    xlnt::python_streambuf buffer(pyfile);
    std::istream stream(&buffer);
    auto table = xlnt::xlsx2arrow(stream);

    return arrow::py::wrap_table(table);
}

PyObject *arrow2xlsx(PyObject *pytable, PyObject *pyfile)
{
    if (!import_pyarrow())
    {
        Py_RETURN_NONE;
    }

    (void)pytable;
    (void)pyfile;
    /*
    auto table = arrow::py::unwrap_table(pytable);
    xlnt::python_streambuf buffer(pyfile);
    std::ostream stream(&buffer);
    xlnt::arrow2xlsx(table, stream);
    */

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
    static const char *keywords[] = { "file", NULL };
    static auto keywords_nc = const_cast<char **>(keywords);

    PyObject *file = NULL;

    if (!PyArg_ParseTupleAndKeywords(args, kwargs, "O", keywords_nc, &file))
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
    static const char *keywords[] = { "table", "file", NULL };
    static auto keywords_nc = const_cast<char **>(keywords);

    PyObject *table = NULL;
    PyObject *file = NULL;

    if (!PyArg_ParseTupleAndKeywords(args, kwargs, "OO", keywords_nc, &table, &file))
    {
        return NULL;
    }

    return arrow2xlsx(table, file);
}

// 2.7/3 compatible based on https://docs.python.org/3/howto/cporting.html

static PyMethodDef xlntpyarrow_methods[] =
{
    { "xlsx2arrow", (PyCFunction)xlntpyarrow_xlsx2arrow, METH_VARARGS | METH_KEYWORDS, xlntpyarrow_xlsx2arrow_doc },
    { "arrow2xlsx", (PyCFunction)xlntpyarrow_arrow2xlsx, METH_VARARGS | METH_KEYWORDS, xlntpyarrow_arrow2xlsx_doc },
    { NULL, NULL, 0, NULL }
};

#if PY_MAJOR_VERSION >= 3

PyDoc_STRVAR(xlntpyarrow_doc, "The xlntpyarrow module");

static PyModuleDef xlntpyarrow_def =
{
    PyModuleDef_HEAD_INIT, // m_base
    "xlntpyarrow", // m_name
    xlntpyarrow_doc, // m_doc
    0, // m_size
    xlntpyarrow_methods, // m_methods
    NULL, // m_slots
    NULL, // m_traverse
    NULL, // m_clear
    NULL, // m_free
};

PyMODINIT_FUNC
PyInit_xlntpyarrow(void)
#else
void
initxlntpyarrow(void)
#endif
{
    PyObject *module = NULL;

#if PY_MAJOR_VERSION >= 3
    module = PyModule_Create(&xlntpyarrow_def);
    return module;
#else
    module = Py_InitModule("xlntpyarrow", xlntpyarrow_methods);
    return;
#endif
}

} // extern "C"
