// Copyright (c) 2014-2017 Thomas Fussell
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
class XLNT_API exception : public std::runtime_error
{
public:
    /// <summary>
    ///
    /// </summary>
    exception(const std::string &message);

    /// <summary>
    ///
    /// </summary>
    exception(const exception &) = default;

    /// <summary>
    ///
    /// </summary>
    virtual ~exception();

    /// <summary>
    ///
    /// </summary>
    void message(const std::string &message);

private:
    /// <summary>
    ///
    /// </summary>
    std::string message_;
};

/// <summary>
/// Exception for a bad parameter value
/// </summary>
class XLNT_API invalid_parameter : public exception
{
public:
    /// <summary>
    ///
    /// </summary>
    invalid_parameter();

    /// <summary>
    ///
    /// </summary>
    invalid_parameter(const invalid_parameter &) = default;

    /// <summary>
    ///
    /// </summary>
    virtual ~invalid_parameter();
};

/// <summary>
/// Exception for bad sheet names.
/// </summary>
class XLNT_API invalid_sheet_title : public exception
{
public:
    /// <summary>
    ///
    /// </summary>
    invalid_sheet_title(const std::string &title);

    /// <summary>
    ///
    /// </summary>
    invalid_sheet_title(const invalid_sheet_title &) = default;

    /// <summary>
    ///
    /// </summary>
    virtual ~invalid_sheet_title();
};

/// <summary>
/// Exception when a referenced number format is not in the stylesheet.
/// </summary>
class XLNT_API missing_number_format : public exception
{
public:
    /// <summary>
    ///
    /// </summary>
    missing_number_format();

    /// <summary>
    ///
    /// </summary>
    virtual ~missing_number_format();
};

/// <summary>
/// Exception for trying to open a non-XLSX file.
/// </summary>
class XLNT_API invalid_file : public exception
{
public:
    /// <summary>
    ///
    /// </summary>
    invalid_file(const std::string &filename);

    /// <summary>
    ///
    /// </summary>
    invalid_file(const invalid_file &) = default;

    /// <summary>
    ///
    /// </summary>
    virtual ~invalid_file();
};

/// <summary>
/// The data submitted which cannot be used directly in Excel files. It
/// must be removed or escaped.
/// </summary>
class XLNT_API illegal_character : public exception
{
public:
    /// <summary>
    ///
    /// </summary>
    illegal_character(char c);

    /// <summary>
    ///
    /// </summary>
    illegal_character(const illegal_character &) = default;

    /// <summary>
    ///
    /// </summary>
    virtual ~illegal_character();
};

/// <summary>
/// Exception for any data type inconsistencies.
/// </summary>
class XLNT_API invalid_data_type : public exception
{
public:
    /// <summary>
    ///
    /// </summary>
    invalid_data_type();

    /// <summary>
    ///
    /// </summary>
    invalid_data_type(const invalid_data_type &) = default;

    /// <summary>
    ///
    /// </summary>
    virtual ~invalid_data_type();
};

/// <summary>
/// Exception for bad column names in A1-style cell references.
/// </summary>
class XLNT_API invalid_column_string_index : public exception
{
public:
    /// <summary>
    ///
    /// </summary>
    invalid_column_string_index();

    /// <summary>
    ///
    /// </summary>
    invalid_column_string_index(const invalid_column_string_index &) = default;

    /// <summary>
    ///
    /// </summary>
    virtual ~invalid_column_string_index();
};

/// <summary>
/// Exception for converting between numeric and A1-style cell references.
/// </summary>
class XLNT_API invalid_cell_reference : public exception
{
public:
    /// <summary>
    ///
    /// </summary>
    invalid_cell_reference(column_t column, row_t row);

    /// <summary>
    ///
    /// </summary>
    invalid_cell_reference(const std::string &reference_string);

    /// <summary>
    ///
    /// </summary>
    invalid_cell_reference(const invalid_cell_reference &) = default;

    /// <summary>
    ///
    /// </summary>
    virtual ~invalid_cell_reference();
};

/// <summary>
/// Exception when setting a class's attribute to an invalid value
/// </summary>
class XLNT_API invalid_attribute : public exception
{
public:
    /// <summary>
    ///
    /// </summary>
    invalid_attribute();

    /// <summary>
    ///
    /// </summary>
    invalid_attribute(const invalid_attribute &) = default;

    /// <summary>
    ///
    /// </summary>
    virtual ~invalid_attribute();
};

/// <summary>
/// Exception for a key that doesn't exist in a container
/// </summary>
class XLNT_API key_not_found : public exception
{
public:
    /// <summary>
    ///
    /// </summary>
    key_not_found();

    /// <summary>
    ///
    /// </summary>
    key_not_found(const key_not_found &) = default;

    /// <summary>
    ///
    /// </summary>
    virtual ~key_not_found();
};

/// <summary>
/// Exception for a workbook with no visible worksheets
/// </summary>
class XLNT_API no_visible_worksheets : public exception
{
public:
    /// <summary>
    ///
    /// </summary>
    no_visible_worksheets();

    /// <summary>
    ///
    /// </summary>
    no_visible_worksheets(const no_visible_worksheets &) = default;

    /// <summary>
    ///
    /// </summary>
    virtual ~no_visible_worksheets();
};

/// <summary>
/// Debug exception for a switch that fell through to the default case
/// </summary>
class XLNT_API unhandled_switch_case : public exception
{
public:
    /// <summary>
    ///
    /// </summary>
    unhandled_switch_case();

    /// <summary>
    ///
    /// </summary>
    virtual ~unhandled_switch_case();
};

/// <summary>
/// Exception for attempting to use a feature which is not supported
/// </summary>
class XLNT_API unsupported : public exception
{
public:
    /// <summary>
    ///
    /// </summary>
    unsupported(const std::string &message);

    /// <summary>
    ///
    /// </summary>
    unsupported(const unsupported &) = default;

    /// <summary>
    ///
    /// </summary>
    virtual ~unsupported();
};

} // namespace xlnt
