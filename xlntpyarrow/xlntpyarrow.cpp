#include <iostream>
#include <memory>
#include <vector>

#include <arrow/api.h>
#include <Python.h> // must be included after Arrow

#include <python_streambuf.hpp>
#include <xlnt/utils/xlntarrow.hpp>

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
    PyObject *file = NULL;
    static const char *keywords[] = { "file", NULL };
    static auto keywords_nc = const_cast<char **>(keywords);

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
    PyObject *obj = NULL;
    static const char *keywords[] = { "file", NULL };
    static auto keywords_nc = const_cast<char **>(keywords);

    if (!PyArg_ParseTupleAndKeywords(args, kwargs, "Oi", keywords_nc, &obj))
    {
        return NULL;
    }

    Py_RETURN_NONE;
}

// 2.7/3 compatible based on https://docs.python.org/3/howto/cporting.html

struct module_state
{
    PyObject *error;
};

#if PY_MAJOR_VERSION >= 3
#define GETSTATE(m) ((struct module_state*)PyModule_GetState(m))
#else
#define GETSTATE(m) (&_state)
static struct module_state _state;
#endif

static PyObject * error_out(PyObject *m)
{
    struct module_state *st = GETSTATE(m);
    PyErr_SetString(st->error, "something bad happened");

    return NULL;
}

static PyMethodDef xlntpyarrow_methods[] =
{
    { "xlsx2arrow", (PyCFunction)xlntpyarrow_xlsx2arrow, METH_VARARGS | METH_KEYWORDS, xlntpyarrow_xlsx2arrow_doc },
    { "arrow2xlsx", (PyCFunction)xlntpyarrow_arrow2xlsx, METH_VARARGS | METH_KEYWORDS, xlntpyarrow_arrow2xlsx_doc },
    { nullptr, NULL, 0, NULL }
};

#if PY_MAJOR_VERSION >= 3

static int xlntpyarrow_traverse(PyObject *m, visitproc visit, void *arg)
{
    Py_VISIT(GETSTATE(m)->error);
    return 0;
}

static int xlntpyarrow_clear(PyObject *m)
{
    Py_CLEAR(GETSTATE(m)->error);
    return 0;
}

PyDoc_STRVAR(xlntpyarrow_doc, "The xlntpyarrow module");

static PyModuleDef xlntpyarrow_def =
{
    PyModuleDef_HEAD_INIT, // m_base
    "xlntpyarrow", // m_name
    xlntpyarrow_doc, // m_doc
    0, // m_size
    xlntpyarrow_methods, // m_methods
    NULL, // m_slots
    xlntpyarrow_traverse, // m_traverse
    xlntpyarrow_clear, // m_clear
    NULL, // m_free
};

#define INITERROR return NULL

PyMODINIT_FUNC
PyInit_xlntpyarrow(void)

#else
#define INITERROR return

void
initmyextension(void)
#endif
{
#if PY_MAJOR_VERSION >= 3
    PyObject *module = PyModule_Create(&xlntpyarrow_def);
#else
    PyObject *module = Py_InitModule("xlntpyarrow", xlntpyarrow_methods);
#endif

    if (module == NULL)
    {
        INITERROR;
    }

    struct module_state *st = GETSTATE(module);

    st->error = PyErr_NewException("xlntpyarrow.Error", NULL, NULL);

    if (st->error == NULL)
    {
        Py_DECREF(module);
        INITERROR;
    }

#if PY_MAJOR_VERSION >= 3
    return module;
#endif
}

} // extern "C"
