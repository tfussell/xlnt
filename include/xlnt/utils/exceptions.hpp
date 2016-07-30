// Copyright (c) 2014-2016 Thomas Fussell
// Copyright (c) 2010-2015 openpyxl
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
#pragma once

#include <cstdint>
#include <stdexcept>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/cell/index_types.hpp>

namespace xlnt {

/// <summary>
/// Parent type of all custom exceptions thrown in this library.
/// </summary>
class XLNT_CLASS exception : public std::runtime_error
{
public:
    exception(const std::string &message);

    void set_message(const std::string &message);

private:
    std::string message_;
};

/// <summary>
/// Exception for a bad parameter value
/// </summary>
class XLNT_CLASS invalid_parameter : public exception
{
public:
    invalid_parameter();
};

/// <summary>
/// Exception for bad sheet names.
/// </summary>
class XLNT_CLASS invalid_sheet_title : public exception
{
public:
    invalid_sheet_title(const std::string &title);
};

/// <summary>
/// Exception when a referenced number format is not in the stylesheet.
/// </summary>
class XLNT_CLASS missing_number_format : public exception
{
public:
    missing_number_format();
};

/// <summary>
/// Exception for trying to open a non-XLSX file.
/// </summary>
class XLNT_CLASS invalid_file : public exception
{
public:
    invalid_file(const std::string &filename);
};

/// <summary>
/// The data submitted which cannot be used directly in Excel files. It
/// must be removed or escaped.
/// </summary>
class XLNT_CLASS illegal_character : public exception
{
public:
    illegal_character(char c);
};

/// <summary>
/// Exception for any data type inconsistencies.
/// </summary>
class XLNT_CLASS invalid_data_type : public exception
{
public:
    invalid_data_type();
};

/// <summary>
/// Exception for bad column names in A1-style cell references.
/// </summary>
class XLNT_CLASS invalid_column_string_index : public exception
{
public:
    invalid_column_string_index();
};

/// <summary>
/// Exception for converting between numeric and A1-style cell references.
/// </summary>
class XLNT_CLASS invalid_cell_reference : public exception
{
public:
    invalid_cell_reference(column_t column, row_t row);
    invalid_cell_reference(const std::string &reference_string);
};

/// <summary>
/// Exception when setting a class's attribute to an invalid value
/// </summary>
class XLNT_CLASS invalid_attribute : public exception
{
public:
    invalid_attribute();
};

/// <summary>
/// Exception for a key that doesn't exist in a container
/// </summary>
class XLNT_CLASS key_not_found : public exception
{
public:
    key_not_found();
};

/// <summary>
/// Exception for a workbook with no visible worksheets
/// </summary>
class XLNT_CLASS no_visible_worksheets : public exception
{
public:
    no_visible_worksheets();
};

} // namespace xlnt
