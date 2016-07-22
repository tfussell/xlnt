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
class XLNT_CLASS error : public std::runtime_error
{
public:
    error();
    error(const std::string &message);

    void set_message(const std::string &message);

private:
    std::string message_;
};

class XLNT_CLASS value_error : public error
{
public:
    value_error();
};

/// <summary>
/// Error for string encoding not matching workbook encoding
/// </summary>
class XLNT_CLASS unicode_decode_error : public error
{
public:
    unicode_decode_error();
    unicode_decode_error(char c);
    unicode_decode_error(std::uint8_t b);
};

/// <summary>
/// Error for bad sheet names.
/// </summary>
class XLNT_CLASS sheet_title_error : public error
{
public:
    sheet_title_error(const std::string &title);
};

/// <summary>
/// Error for trying to modify a read-only workbook.
/// </summary>
class XLNT_CLASS read_only_workbook_error : public error
{
public:
    read_only_workbook_error();
};

/// <summary>
/// Error for incorrectly formatted named ranges.
/// </summary>
class XLNT_CLASS named_range_error : public error
{
public:
    named_range_error();
};

/// <summary>
/// Error when a referenced number format is not in the stylesheet.
/// </summary>
class XLNT_CLASS missing_number_format : public error
{
public:
    missing_number_format();
};

/// <summary>
/// Error for trying to open a non-OOXML file.
/// </summary>
class XLNT_CLASS invalid_file_error : public error
{
public:
    invalid_file_error(const std::string &filename);
};

/// <summary>
/// The data submitted which cannot be used directly in Excel files. It
/// must be removed or escaped.
/// </summary>
class XLNT_CLASS illegal_character_error : public error
{
public:
    illegal_character_error(char c);
};

/// <summary>
/// Error for any data type inconsistencies.
/// </summary>
class XLNT_CLASS data_type_error : public error
{
public:
    data_type_error();
};

/// <summary>
/// Error for bad column names in A1-style cell references.
/// </summary>
class XLNT_CLASS column_string_index_error : public error
{
public:
    column_string_index_error();
};

/// <summary>
/// Error for converting between numeric and A1-style cell references.
/// </summary>
class XLNT_CLASS cell_coordinates_error : public error
{
public:
    cell_coordinates_error(column_t column, row_t row);
    cell_coordinates_error(const std::string &coord_string);
};

/// <summary>
/// Error when an attribute value is invalid.
/// </summary>
class XLNT_CLASS attribute_error : public error
{
public:
    attribute_error();
};

/// <summary>
/// key_error
/// </summary>
class XLNT_CLASS key_error : public error
{
public:
    key_error();
};

} // namespace xlnt
