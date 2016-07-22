:mod:`xlnt` - user-friendly xlsx library for C++14
===========================================================================

.. module:: xlnt
.. moduleauthor:: Thomas Fussell

:Author: Thomas Fussell
:Source code: https://github.com/tfussell/xlnt
:Issues: https://github.com/tfussell/xlnt/issues
:Generated: |today|
:License: MIT
:Version: |release|

Introduction
------------

xlnt is a C++14 library for reading, writing, and modifying XLSX files as described in [ECMA 376](http://www.ecma-international.org/publications/standards/Ecma-376.htm). The API is based on [openpyxl](https://bitbucket.org/openpyxl/openpyxl), a Python library to read/write Excel 2007 xlsx/xlsm files, and ultimately on [PHPExcel](https://github.com/PHPOffice/PHPExcel), pure PHP library for reading and writing spreadsheet files upon which openpyxl was based. This project is still very much a work in progress, but the core development work is complete.

Support
+++++++

Sample code:
++++++++++++

.. literalinclude:: /samples/sample.cpp


How to Contribute Code
----------------------

See :ref:`development`

Installation
------------

Getting the source
------------------

Usage examples
--------------

Tutorial
++++++++

.. toctree::

    tutorial

Cookbook
++++++++

.. toctree::

    cookbook

Charts
++++++

.. toctree::

    charts/introduction

Comments
++++++++

.. toctree::

    comments

Working with styles
+++++++++++++++++++

.. toctree::

    styles

Conditional Formatting
++++++++++++++++++++++

.. toctree::

    formatting

Data Validation
+++++++++++++++

.. toctree::

    validation

Parsing Formulas
++++++++++++++++

.. toctree::

    formula


Information for Developers
--------------------------

.. toctree::

    development
    windows-development

API Documentation
------------------

.. toctree::
    :maxdepth: 2

    api/xlnt


Indices and tables
==================

* :ref:`genindex`
* :ref:`modindex`
* :ref:`search`


Release Notes
=============

.. toctree::
    :maxdepth: 1

    changes
