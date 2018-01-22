// Copyright (c) 2014-2018 Thomas Fussell
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

#include <cctype>

#include <xlnt/cell/cell_reference.hpp>
#include <xlnt/utils/exceptions.hpp>
#include <xlnt/worksheet/range_reference.hpp>

#include <detail/constants.hpp>

namespace xlnt {

std::size_t cell_reference_hash::operator()(const cell_reference &k) const
{
    return k.row() * constants::max_column().index + k.column_index();
}

cell_reference &cell_reference::make_absolute(bool absolute_column, bool absolute_row)
{
    column_absolute(absolute_column);
    row_absolute(absolute_row);

    return *this;
}

cell_reference::cell_reference()
    : cell_reference(1, 1)
{
}

cell_reference::cell_reference(const std::string &string)
{
    auto split = split_reference(string, absolute_column_, absolute_row_);

    column(split.first);
    row(split.second);
}

cell_reference::cell_reference(const char *reference_string)
    : cell_reference(std::string(reference_string))
{
}

cell_reference::cell_reference(column_t column_index, row_t row)
    : column_(column_index), row_(row), absolute_row_(false), absolute_column_(false)
{
    if (row_ == 0
      || column_ == 0
      || !(row_ <= constants::max_row())
      || !(column_ <= constants::max_column()))
    {
        throw invalid_cell_reference(column_, row_);
    }
}

range_reference cell_reference::operator,(const xlnt::cell_reference &other) const
{
    return range_reference(*this, other);
}

std::string cell_reference::to_string() const
{
    std::string string_representation;

    if (absolute_column_)
    {
        string_representation.append("$");
    }

    string_representation.append(column_.column_string());

    if (absolute_row_)
    {
        string_representation.append("$");
    }

    string_representation.append(std::to_string(row_));

    return string_representation;
}

range_reference cell_reference::to_range() const
{
    return range_reference(column_, row_, column_, row_);
}

std::pair<std::string, row_t> cell_reference::split_reference(const std::string &reference_string)
{
    bool ignore1, ignore2;
    return split_reference(reference_string, ignore1, ignore2);
}

std::pair<std::string, row_t> cell_reference::split_reference(
    const std::string &reference_string, bool &absolute_column, bool &absolute_row)
{
    absolute_column = false;
    absolute_row = false;

    // Convert a coordinate string like 'B12' to a tuple ('B', 12)
    bool column_part = true;

    std::string column_string;

    for (auto character : reference_string)
    {
        auto upper = static_cast<char>(std::toupper(static_cast<std::uint8_t>(character)));

        if (std::isalpha(character))
        {
            if (column_part)
            {
                column_string.append(1, upper);
            }
            else
            {
                throw invalid_cell_reference(reference_string);
            }
        }
        else if (character == '$')
        {
            if (column_part)
            {
                if (column_string.empty())
                {
                    column_string.append(1, upper);
                }
                else
                {
                    column_part = false;
                }
            }
        }
        else
        {
            if (column_part)
            {
                column_part = false;
            }
            else if (!std::isdigit(character))
            {
                throw invalid_cell_reference(reference_string);
            }
        }
    }

    std::string row_string = reference_string.substr(column_string.length());

    if (row_string.length() == 0)
    {
        throw invalid_cell_reference(reference_string);
    }

    if (column_string[0] == '$')
    {
        absolute_row = true;
        column_string = column_string.substr(1);
    }

    if (row_string[0] == '$')
    {
        absolute_column = true;
        row_string = row_string.substr(1);
    }

    return {column_string, std::stoi(row_string)};
}

bool cell_reference::column_absolute() const
{
    return absolute_column_;
}

void cell_reference::column_absolute(bool absolute_column)
{
    absolute_column_ = absolute_column;
}

bool cell_reference::row_absolute() const
{
    return absolute_row_;
}

void cell_reference::row_absolute(bool absolute_row)
{
    absolute_row_ = absolute_row;
}

column_t cell_reference::column() const
{
    return column_;
}

void cell_reference::column(const std::string &column_string)
{
    column_ = column_t(column_string);
}

column_t::index_t cell_reference::column_index() const
{
    return column_.index;
}

void cell_reference::column_index(column_t column)
{
    column_ = column;
}

row_t cell_reference::row() const
{
    return row_;
}

void cell_reference::row(row_t row)
{
    row_ = row;
}

bool cell_reference::operator==(const std::string &reference_string) const
{
    return *this == cell_reference(reference_string);
}

bool cell_reference::operator==(const char *reference_string) const
{
    return *this == std::string(reference_string);
}

bool cell_reference::operator!=(const cell_reference &comparand) const
{
    return !(*this == comparand);
}

bool cell_reference::operator!=(const std::string &reference_string) const
{
    return *this != cell_reference(reference_string);
}

bool cell_reference::operator!=(const char *reference_string) const
{
    return *this != std::string(reference_string);
}

cell_reference cell_reference::make_offset(int column_offset, int row_offset) const
{
    // TODO: check for overflow/underflow
    auto relative_column = static_cast<column_t::index_t>(static_cast<int>(column_.index) + column_offset);
    auto relative_row = static_cast<row_t>(static_cast<int>(row_) + row_offset);

    return cell_reference(relative_column, relative_row);
}

bool cell_reference::operator==(const cell_reference &comparand) const
{
    return comparand.column_ == column_
        && comparand.row_ == row_
        && absolute_column_ == comparand.absolute_column_
        && absolute_row_ == comparand.absolute_row_;
}

} // namespace xlnt
