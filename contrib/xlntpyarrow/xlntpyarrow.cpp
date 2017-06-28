#include <iostream>
#include <Python.h>

class abc {
public:
    static void def()
    {
        std::cout << "abc" << std::endl;
    }
};

extern "C" {

/*
 * Implements an example function.
 */
PyDoc_STRVAR(xlntpyarrow_example_doc, "example(obj, number)\
\
Example function");

PyObject *xlntpyarrow_example(PyObject *self, PyObject *args, PyObject *kwargs) {
    /* Shared references that do not need Py_DECREF before returning. */
    PyObject *obj = NULL;
    int number = 0;

    abc::def();

    /* Parse positional and keyword arguments */
    static char* keywords[] = { "obj", "number", NULL };
    if (!PyArg_ParseTupleAndKeywords(args, kwargs, "Oi", keywords, &obj, &number)) {
        return NULL;
    }

    /* Function implementation starts here */

    if (number < 0) {
        PyErr_SetObject(PyExc_ValueError, obj);
        return NULL;    /* return NULL indicates error */
    }

    Py_RETURN_NONE;
}

/*
 * List of functions to add to xlntpyarrow in exec_xlntpyarrow().
 */
static PyMethodDef xlntpyarrow_functions[] = {
    { "example", (PyCFunction)xlntpyarrow_example, METH_VARARGS | METH_KEYWORDS, xlntpyarrow_example_doc },
    { NULL, NULL, 0, NULL } /* marks end of array */
};

/*
 * Initialize xlntpyarrow. May be called multiple times, so avoid
 * using static state.
 */
int exec_xlntpyarrow(PyObject *module) {
    PyModule_AddFunctions(module, xlntpyarrow_functions);

    PyModule_AddStringConstant(module, "__author__", "Thomas");
    PyModule_AddStringConstant(module, "__version__", "1.0.0");
    PyModule_AddIntConstant(module, "year", 2017);

    return 0; /* success */
}

/*
 * Documentation for xlntpyarrow.
 */
PyDoc_STRVAR(xlntpyarrow_doc, "The xlntpyarrow module");


static PyModuleDef_Slot xlntpyarrow_slots[] = {
    { Py_mod_exec, exec_xlntpyarrow },
    { 0, NULL }
};

static PyModuleDef xlntpyarrow_def = {
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

PyMODINIT_FUNC PyInit_xlntpyarrow() {
    return PyModuleDef_Init(&xlntpyarrow_def);
}

}