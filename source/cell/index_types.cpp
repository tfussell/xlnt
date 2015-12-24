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
#include <locale>

#include <xlnt/cell/index_types.hpp>
#include <xlnt/utils/exceptions.hpp>

#include <detail/constants.hpp>

namespace xlnt {

column_t::index_t column_t::column_index_from_string(const std::string &column_string)
{
    if (column_string.length() > 3 || column_string.empty())
    {
        throw column_string_index_exception();
    }

    column_t::index_t column_index = 0;
    int place = 1;

    for (int i = static_cast<int>(column_string.length()) - 1; i >= 0; i--)
    {
        if (!std::isalpha(column_string[static_cast<std::size_t>(i)], std::locale::classic()))
        {
            throw column_string_index_exception();
        }

        auto char_index = std::toupper(column_string[static_cast<std::size_t>(i)], std::locale::classic()) - 'A';

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
    if (column_index < constants::MinColumn() || column_index > constants::MaxColumn())
    {
        //        auto msg = "Column index out of bounds: " + std::to_string(column_index);
        throw column_string_index_exception();
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


/// <summary>
/// Default column_t is the first (left-most) column.
/// </summary>
column_t::column_t() : index(1) {}

/// <summary>
/// Construct a column from a number.
/// </summary>
column_t::column_t(index_t column_index) : index(column_index) {}

/// <summary>
/// Construct a column from a string.
/// </summary>
column_t::column_t(const std::string &column_string) : index(column_index_from_string(column_string)) {}

/// <summary>
/// Construct a column from a string.
/// </summary>
column_t::column_t(const char *column_string) : column_t(std::string(column_string)) {}

/// <summary>
/// Copy constructor
/// </summary>
column_t::column_t(const column_t &other) : column_t(other.index) {}

/// <summary>
/// Move constructor
/// </summary>
column_t::column_t(column_t &&other) { swap(*this, other); }

/// <summary>
/// Return a string representation of this column index.
/// </summary>
std::string column_t::column_string() const { return column_string_from_index(index); }

/// <summary>
/// Set this column to be the same as rhs's and return reference to self.
/// </summary>
column_t &column_t::operator=(column_t rhs) { swap(*this, rhs); return *this; }

/// <summary>
/// Set this column to be equal to rhs and return reference to self.
/// </summary>
column_t &column_t::operator=(const std::string &rhs) { return *this = column_t(rhs); }

/// <summary>
/// Set this column to be equal to rhs and return reference to self.
/// </summary>
column_t &column_t::operator=(const char *rhs) { return *this = column_t(rhs); }

/// <summary>
/// Return true if this column refers to the same column as other.
/// </summary>
bool column_t::operator==(const column_t &other) const { return index == other.index; }

/// <summary>
/// Return true if this column doesn't refer to the same column as other.
/// </summary>
bool column_t::operator!=(const column_t &other) const { return !(*this == other); }

/// <summary>
/// Return true if this column refers to the same column as other.
/// </summary>
bool column_t::operator==(int other) const { return *this == column_t(other); }

/// <summary>
/// Return true if this column refers to the same column as other.
/// </summary>
bool column_t::operator==(index_t other) const { return *this == column_t(other); }

/// <summary>
/// Return true if this column refers to the same column as other.
/// </summary>
bool column_t::operator==(const std::string &other) const { return *this == column_t(other); }

/// <summary>
/// Return true if this column refers to the same column as other.
/// </summary>
bool column_t::operator==(const char *other) const { return *this == column_t(other); }

/// <summary>
/// Return true if this column doesn't refer to the same column as other.
/// </summary>
bool column_t::operator!=(int other) const { return !(*this == other); }

/// <summary>
/// Return true if this column doesn't refer to the same column as other.
/// </summary>
bool column_t::operator!=(index_t other) const { return !(*this == other); }

/// <summary>
/// Return true if this column doesn't refer to the same column as other.
/// </summary>
bool column_t::operator!=(const std::string &other) const { return !(*this == other); }

/// <summary>
/// Return true if this column doesn't refer to the same column as other.
/// </summary>
bool column_t::operator!=(const char *other) const { return !(*this == other); }

/// <summary>
/// Return true if other is to the right of this column.
/// </summary>
bool column_t::operator>(const column_t &other) const { return index > other.index; }

/// <summary>
/// Return true if other is to the right of or equal to this column.
/// </summary>
bool column_t::operator>=(const column_t &other) const { return index >= other.index; }

/// <summary>
/// Return true if other is to the left of this column.
/// </summary>
bool column_t::operator<(const column_t &other) const { return index < other.index; }

/// <summary>
/// Return true if other is to the left of or equal to this column.
/// </summary>
bool column_t::operator<=(const column_t &other) const { return index <= other.index; }

/// <summary>
/// Return true if other is to the right of this column.
/// </summary>
bool column_t::operator>(const column_t::index_t &other) const { return index > other; }

/// <summary>
/// Return true if other is to the right of or equal to this column.
/// </summary>
bool column_t::operator>=(const column_t::index_t &other) const { return index >= other; }

/// <summary>
/// Return true if other is to the left of this column.
/// </summary>
bool column_t::operator<(const column_t::index_t &other) const { return index < other; }

/// <summary>
/// Return true if other is to the left of or equal to this column.
/// </summary>
bool column_t::operator<=(const column_t::index_t &other) const { return index <= other; }

/// <summary>
/// Pre-increment this column, making it point to the column one to the right.
/// </summary>
column_t &column_t::operator++() { index++; return *this; }

/// <summary>
/// Pre-deccrement this column, making it point to the column one to the left.
/// </summary>
column_t &column_t::operator--() { index--; return *this; }

/// <summary>
/// Post-increment this column, making it point to the column one to the right and returning the old column.
/// </summary>
column_t column_t::operator++(int) { column_t copy(index); ++(*this); return copy; }

/// <summary>
/// Post-decrement this column, making it point to the column one to the left and returning the old column.
/// </summary>
column_t column_t::operator--(int) { column_t copy(index); --(*this); return copy; }

/// <summary>
/// Return the result of adding rhs to this column.
/// </summary>
column_t column_t::operator+(const column_t &rhs) { column_t copy(*this); copy.index += rhs.index; return copy; }

/// <summary>
/// Return the result of adding rhs to this column.
/// </summary>
column_t column_t::operator-(const column_t &rhs) { column_t copy(*this); copy.index -= rhs.index; return copy; }

/// <summary>
/// Return the result of adding rhs to this column.
/// </summary>
column_t column_t::operator*(const column_t &rhs) { column_t copy(*this); copy.index *= rhs.index; return copy; }

/// <summary>
/// Return the result of adding rhs to this column.
/// </summary>
column_t column_t::operator/(const column_t &rhs) { column_t copy(*this); copy.index /= rhs.index; return copy; }

/// <summary>
/// Return the result of adding rhs to this column.
/// </summary>
column_t column_t::operator%(const column_t &rhs) { column_t copy(*this); copy.index %= rhs.index; return copy; }

/// <summary>
/// Add rhs to this column and return a reference to this column.
/// </summary>
column_t &column_t::operator+=(const column_t &rhs) { return *this = (*this + rhs); }

/// <summary>
/// Subtrac rhs from this column and return a reference to this column.
/// </summary>
column_t &column_t::operator-=(const column_t &rhs) { return *this = (*this - rhs); }

/// <summary>
/// Multiply this column by rhs and return a reference to this column.
/// </summary>
column_t &column_t::operator*=(const column_t &rhs) { return *this = (*this * rhs); }

/// <summary>
/// Divide this column by rhs and return a reference to this column.
/// </summary>
column_t &column_t::operator/=(const column_t &rhs) { return *this = (*this / rhs); }

/// <summary>
/// Mod this column by rhs and return a reference to this column.
/// </summary>
column_t &column_t::operator%=(const column_t &rhs) { return *this = (*this % rhs); }

/// <summary>
/// Return true if other is to the right of this column.
/// </summary>
bool operator>(const column_t::index_t &left, const column_t &right) { return column_t(left) > right; }

/// <summary>
/// Return true if other is to the right of or equal to this column.
/// </summary>
bool operator>=(const column_t::index_t &left, const column_t &right) { return column_t(left) >= right; }

/// <summary>
/// Return true if other is to the left of this column.
/// </summary>
bool operator<(const column_t::index_t &left, const column_t &right) { return column_t(left) < right; }

/// <summary>
/// Return true if other is to the left of or equal to this column.
/// </summary>
bool operator<=(const column_t::index_t &left, const column_t &right) { return column_t(left) <= right; }

/// <summary>
/// Swap the columns that left and right refer to.
/// </summary>
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
