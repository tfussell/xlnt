// Copyright (c) 2015 Thomas Fussell
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

#include <stdexcept>
#include <string>

#include <xlnt/cell/types.hpp>

#include <xlnt/xlnt_config.hpp>

namespace xlnt {

/// <summary>
/// Error for converting between numeric and A1-style cell references.
/// </summary>
class XLNT_CLASS cell_coordinates_exception
{
  public:
    cell_coordinates_exception(column_t column, row_t row);
    cell_coordinates_exception(const string &coord_string);
};

/// <summary>
/// The data submitted which cannot be used directly in Excel files. It
/// must be removed or escaped.
/// </summary>
class XLNT_CLASS illegal_character_error
{
  public:
    illegal_character_error(utf32_char c);
};

/// <summary>
/// Error for bad column names in A1-style cell references.
/// </summary>
class XLNT_CLASS column_string_index_exception
{
  public:
    column_string_index_exception();
};

/// <summary>
/// Error for any data type inconsistencies.
/// </summary>
class XLNT_CLASS data_type_exception
{
  public:
    data_type_exception();
};

/// <summary>
/// Error for badly formatted named ranges.
/// </summary>
class XLNT_CLASS named_range_exception
{
  public:
    named_range_exception();
};

/// <summary>
/// Error for bad sheet names.
/// </summary>
class XLNT_CLASS sheet_title_exception
{
  public:
    sheet_title_exception(const string &title);
};

/// <summary>
/// Error for trying to open a non-ooxml file.
/// </summary>
class XLNT_CLASS invalid_file_exception
{
  public:
    invalid_file_exception(const string &filename);
};

/// <summary>
/// Error for trying to modify a read-only workbook.
/// </summary>
class XLNT_CLASS read_only_workbook_exception
{
  public:
    read_only_workbook_exception();
};

/// <summary>
/// Error when a references number format is not in the stylesheet.
/// </summary>
class XLNT_CLASS missing_number_format
{
  public:
    missing_number_format();
};

/// <summary>
/// Error when an attribute value is invalid.
/// </summary>
class XLNT_CLASS attribute_error
{
  public:
    attribute_error();
};

class XLNT_CLASS value_error
{
  public:
    value_error()
    {
    }
};

} // namespace xlnt
