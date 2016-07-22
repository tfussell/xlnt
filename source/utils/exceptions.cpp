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

#include <xlnt/utils/exceptions.hpp>

namespace xlnt {

error::error() : std::runtime_error("xlnt::error")
{
}

error::error(const std::string &message)
    : std::runtime_error("xlnt::error : " + message)
{
    set_message(message);
}

void error::set_message(const std::string &message)
{
    message_ = message;
}

sheet_title_error::sheet_title_error(const std::string &title)
    : error(std::string("bad worksheet title: ") + title)
{
}

column_string_index_error::column_string_index_error()
    : error("column string index error")
{
}

data_type_error::data_type_error()
    : error("data type error")
{
}

named_range_error::named_range_error()
    : error("named range not found or not owned by this worksheet")
{
}

invalid_file_error::invalid_file_error(const std::string &filename)
    : error(std::string("couldn't open file: (") + filename + ")")
{
}

cell_coordinates_error::cell_coordinates_error(column_t column, row_t row)
    : error(std::string("bad cell coordinates: (") + std::to_string(column.index) + ", " + std::to_string(row) +
                         ")")
{
}

cell_coordinates_error::cell_coordinates_error(const std::string &coord_string)
    : error(std::string("bad cell coordinates: (") + (coord_string.empty() ? "<empty>" : coord_string) + ")")
{
}

illegal_character_error::illegal_character_error(char c)
    : error(std::string("illegal character: (") + std::to_string(static_cast<unsigned char>(c)) + ")")
{
}

unicode_decode_error::unicode_decode_error()
    : error("unicode decode error")
{
}

unicode_decode_error::unicode_decode_error(char)
    : error("unicode decode error")
{
}

value_error::value_error()
    : error("value error")
{
}

read_only_workbook_error::read_only_workbook_error()
    : error("workbook is read-only")
{
}

attribute_error::attribute_error()
    : error("bad attribute")
{
}

key_error::key_error()
    : error("key not found in container")
{
}

} // namespace xlnt
