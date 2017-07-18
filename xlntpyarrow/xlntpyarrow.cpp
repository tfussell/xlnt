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

#include <Python.h>
#include <methods.hpp>

extern "C" {

PyDoc_STRVAR(xlntpyarrow_xlsx2arrow_doc, "xlsx2arrow(in_file)\
\
Returns an arrow table representing the given XLSX file object.");

PyDoc_STRVAR(xlntpyarrow_arrow2xlsx_doc, "arrow2xlsx(table, out_file)\
\
Writes the given arrow table to out_file as an XLSX file.");

// 2.7/3 compatible based on https://docs.python.org/3/howto/cporting.html

static PyMethodDef xlntpyarrow_methods[] =
{
    { "xlsx2arrow", (PyCFunction)xlntpyarrow_xlsx2arrow,
        METH_VARARGS | METH_KEYWORDS, xlntpyarrow_xlsx2arrow_doc },
    { "arrow2xlsx", (PyCFunction)xlntpyarrow_arrow2xlsx,
        METH_VARARGS | METH_KEYWORDS, xlntpyarrow_arrow2xlsx_doc },
    { nullptr, nullptr, 0, nullptr }
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
    nullptr, // m_slots
    nullptr, // m_traverse
    nullptr, // m_clear
    nullptr, // m_free
};

PyMODINIT_FUNC
PyInit_xlntpyarrow(void)
#else
void
initxlntpyarrow(void)
#endif
{
    PyObject *module = nullptr;

#if PY_MAJOR_VERSION >= 3
    module = PyModule_Create(&xlntpyarrow_def);
    return module;
#else
    module = Py_InitModule("xlntpyarrow", xlntpyarrow_methods);
    return;
#endif
}

} // extern "C"
