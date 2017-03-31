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

#include <xlnt/utils/exceptions.hpp>

namespace xlnt {

exception::exception(const std::string &message)
    : std::runtime_error("xlnt::exception : " + message)
{
    this->message(message);
}

exception::~exception()
{
}

void exception::message(const std::string &message)
{
    message_ = message;
}

missing_number_format::~missing_number_format()
{
}

unhandled_switch_case::unhandled_switch_case()
    : xlnt::exception("unhandled switch case")
{
}

unhandled_switch_case::~unhandled_switch_case()
{
}

invalid_sheet_title::invalid_sheet_title(const std::string &title)
    : exception(std::string("bad worksheet title: ") + title)
{
}

invalid_sheet_title::~invalid_sheet_title()
{
}

invalid_column_index::invalid_column_index()
    : exception("column string index error")
{
}

invalid_column_index::~invalid_column_index()
{
}

invalid_data_type::invalid_data_type()
    : exception("data type error")
{
}

invalid_data_type::~invalid_data_type()
{
}

invalid_file::invalid_file(const std::string &filename)
    : exception(std::string("couldn't open file: (") + filename + ")")
{
}

invalid_file::~invalid_file()
{
}

invalid_cell_reference::invalid_cell_reference(column_t column, row_t row)
    : exception(
          std::string("bad cell coordinates: (") + std::to_string(column.index) + ", " + std::to_string(row) + ")")
{
}

invalid_cell_reference::invalid_cell_reference(const std::string &coord_string)
    : exception(std::string("bad cell coordinates: (") + (coord_string.empty() ? "<empty>" : coord_string) + ")")
{
}

invalid_cell_reference::~invalid_cell_reference()
{
}

illegal_character::illegal_character(char c)
    : exception(std::string("illegal character: (") + std::to_string(static_cast<unsigned char>(c)) + ")")
{
}

illegal_character::~illegal_character()
{
}

invalid_parameter::invalid_parameter()
    : exception("invalid parameter")
{
}

invalid_parameter::~invalid_parameter()
{
}

invalid_attribute::invalid_attribute()
    : exception("bad attribute")
{
}

invalid_attribute::~invalid_attribute()
{
}

key_not_found::key_not_found()
    : exception("key not found in container")
{
}

key_not_found::~key_not_found()
{
}

no_visible_worksheets::no_visible_worksheets()
    : exception("workbook needs at least one non-hidden worksheet to be saved")
{
}

no_visible_worksheets::~no_visible_worksheets()
{
}

unsupported::unsupported(const std::string &message)
    : exception(message)
{
}

unsupported::~unsupported()
{
}

} // namespace xlnt
