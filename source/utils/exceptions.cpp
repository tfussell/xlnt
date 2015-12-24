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

sheet_title_exception::sheet_title_exception(const std::string &title)
    : std::runtime_error(std::string("bad worksheet title: ") + title)
{
}

column_string_index_exception::column_string_index_exception()
    : std::runtime_error("column string index exception")
{
}

data_type_exception::data_type_exception()
    : std::runtime_error("data type exception")
{
}

attribute_error::attribute_error()
    : std::runtime_error("attribute error")
{
}

named_range_exception::named_range_exception()
    : std::runtime_error("named range not found or not owned by this worksheet")
{
}

invalid_file_exception::invalid_file_exception(const std::string &filename)
    : std::runtime_error(std::string("couldn't open file: (") + filename + ")")
{
}

cell_coordinates_exception::cell_coordinates_exception(column_t column, row_t row)
    : std::runtime_error(std::string("bad cell coordinates: (") + std::to_string(column.index) + ", " + std::to_string(row) +
                         ")")
{
}

cell_coordinates_exception::cell_coordinates_exception(const std::string &coord_string)
    : std::runtime_error(std::string("bad cell coordinates: (") + (coord_string.empty() ? "<empty>" : coord_string) + ")")
{
}

illegal_character_error::illegal_character_error(char c)
    : std::runtime_error(std::string("illegal character: (") + std::to_string(static_cast<unsigned char>(c)) + ")")
{
}

unicode_decode_error::unicode_decode_error()
    : std::runtime_error("unicode decode error")
{
}

unicode_decode_error::unicode_decode_error(char)
    : std::runtime_error("unicode decode error")
{
}

value_error::value_error()
    : std::runtime_error("value error")
{
}

} // namespace xlnt
