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

#include <xlnt/cell/index_types.hpp>
#include <xlnt/utils/exceptions.hpp>
#include <detail/constants.hpp>

namespace xlnt {

column_t::index_t column_t::column_index_from_string(const std::string &column_string)
{
    if (column_string.length() > 3 || column_string.empty())
    {
        throw invalid_column_index();
    }

    column_t::index_t column_index = 0;
    int place = 1;

    for (int i = static_cast<int>(column_string.length()) - 1; i >= 0; i--)
    {
        if (!std::isalpha(column_string[static_cast<std::size_t>(i)]))
        {
            throw invalid_column_index();
        }

        auto char_index = std::toupper(column_string[static_cast<std::size_t>(i)]) - 'A';

        column_index += static_cast<column_t::index_t>((char_index + 1) * place);
        place *= 26;
    }

    return column_index;
}

// Convert a column number into a column letter (3 -> 'C')
// Right shift the column col_idx by 26 to find column letters in reverse
// order.These numbers are 1 - based, and can be converted to ASCII
// ordinals by adding 64.
std::string column_t::column_string_from_index(column_t::index_t column_index)
{
    // these indicies corrospond to A->ZZZ and include all allowed
    // columns
    if (column_index < constants::min_column() || column_index > constants::max_column())
    {
        throw invalid_column_index();
    }

    int temp = static_cast<int>(column_index);
    std::string column_letter = "";

    while (temp > 0)
    {
        int quotient = temp / 26, remainder = temp % 26;

        // check for exact division and borrow if needed
        if (remainder == 0)
        {
            quotient -= 1;
            remainder = 26;
        }

        column_letter = std::string(1, char(remainder + 64)) + column_letter;
        temp = quotient;
    }

    return column_letter;
}

column_t::column_t()
    : index(1)
{
}

column_t::column_t(index_t column_index)
    : index(column_index)
{
}

column_t::column_t(const std::string &column_string)
    : index(column_index_from_string(column_string))
{
}

column_t::column_t(const char *column_string)
    : column_t(std::string(column_string))
{
}

column_t::column_t(const column_t &other)
    : column_t(other.index)
{
}

column_t::column_t(column_t &&other)
{
    swap(*this, other);
}

std::string column_t::column_string() const
{
    return column_string_from_index(index);
}

column_t &column_t::operator=(column_t rhs)
{
    swap(*this, rhs);
    return *this;
}

column_t &column_t::operator=(const std::string &rhs)
{
    return *this = column_t(rhs);
}

column_t &column_t::operator=(const char *rhs)
{
    return *this = column_t(rhs);
}

bool column_t::operator==(const column_t &other) const
{
    return index == other.index;
}

bool column_t::operator!=(const column_t &other) const
{
    return !(*this == other);
}

bool column_t::operator==(int other) const
{
    return *this == column_t(static_cast<index_t>(other));
}

bool column_t::operator==(index_t other) const
{
    return *this == column_t(other);
}

bool column_t::operator==(const std::string &other) const
{
    return *this == column_t(other);
}

bool column_t::operator==(const char *other) const
{
    return *this == column_t(other);
}

bool column_t::operator!=(int other) const
{
    return !(*this == other);
}

bool column_t::operator!=(index_t other) const
{
    return !(*this == other);
}

bool column_t::operator!=(const std::string &other) const
{
    return !(*this == other);
}

bool column_t::operator!=(const char *other) const
{
    return !(*this == other);
}

bool column_t::operator>(const column_t &other) const
{
    return index > other.index;
}

bool column_t::operator>=(const column_t &other) const
{
    return index >= other.index;
}

bool column_t::operator<(const column_t &other) const
{
    return index < other.index;
}

bool column_t::operator<=(const column_t &other) const
{
    return index <= other.index;
}

bool column_t::operator>(const column_t::index_t &other) const
{
    return index > other;
}

bool column_t::operator>=(const column_t::index_t &other) const
{
    return index >= other;
}

bool column_t::operator<(const column_t::index_t &other) const
{
    return index < other;
}

bool column_t::operator<=(const column_t::index_t &other) const
{
    return index <= other;
}

column_t &column_t::operator++()
{
    index++;
    return *this;
}

column_t &column_t::operator--()
{
    index--;
    return *this;
}

column_t column_t::operator++(int)
{
    column_t copy(index);
    ++(*this);
    return copy;
}

column_t column_t::operator--(int)
{
    column_t copy(index);
    --(*this);
    return copy;
}

column_t operator+(column_t lhs, const column_t &rhs)
{
    lhs += rhs;
    return lhs;
}

column_t operator-(column_t lhs, const column_t &rhs)
{
    lhs -= rhs;
    return lhs;
}

column_t &column_t::operator+=(const column_t &rhs)
{
    index += rhs.index;
    return *this;
}

column_t &column_t::operator-=(const column_t &rhs)
{
    index -= rhs.index;
    return *this;
}

bool operator>(const column_t::index_t &left, const column_t &right)
{
    return column_t(left) > right;
}

bool operator>=(const column_t::index_t &left, const column_t &right)
{
    return column_t(left) >= right;
}

bool operator<(const column_t::index_t &left, const column_t &right)
{
    return column_t(left) < right;
}

bool operator<=(const column_t::index_t &left, const column_t &right)
{
    return column_t(left) <= right;
}

void swap(column_t &left, column_t &right)
{
    using std::swap;
    swap(left.index, right.index);
}

std::size_t column_hash::operator()(const column_t &k) const
{
    static std::hash<column_t::index_t> hasher;
    return hasher(k.index);
}

} // namespace xlnt
