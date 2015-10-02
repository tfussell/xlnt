// Copyright (c) 2015 Thomas Fussell
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

#include <memory>
#include <string>
#include <unordered_map>

#include "../styles/style.hpp"
#include "../common/types.hpp"

namespace xlnt {
    
class cell_reference;
class comment;
class relationship;
class value;
class worksheet;

struct date;
struct datetime;
struct time;
struct timedelta;

namespace detail {    
struct cell_impl;
} // namespace detail

/// <summary>
/// Describes cell associated properties.
/// </summary>
/// <remarks>
/// Properties of interest include style, type, value, and address.
/// The Cell class is required to know its value and type, display options,
/// and any other features of an Excel cell.Utilities for referencing
/// cells using Excel's 'A1' column/row nomenclature are also provided.
/// </remarks>
class cell
{
public:
    static const std::unordered_map<std::string, int> ErrorCodes;
    
    cell();
    cell(worksheet ws, const cell_reference &reference);
    cell(worksheet ws, const cell_reference &reference, const value &initial_value);
    
    std::string get_column() const;
    row_t get_row() const;
    
    std::string to_string() const;   
    
    bool is_merged() const;
    void set_merged(bool merged);
    
    relationship get_hyperlink() const;
    void set_hyperlink(const std::string &value);
    bool has_hyperlink() const;
    
    void set_number_format(const std::string &format_code);
    
    bool has_style() const;
    style &get_style();
    const style &get_style() const;
    void set_style(const style &s);

    std::pair<int, int> get_anchor() const;

    bool garbage_collectible() const;
    
    cell_reference get_reference() const;
    
    bool is_date() const;
    
    comment get_comment() const;
    void set_comment(const comment &comment);
    void clear_comment();
    bool has_comment() const;

    std::string get_formula() const;
    void set_formula(const std::string &formula);
    void clear_formula();
    bool has_formula() const;

    std::string get_error() const;
    void set_error(const std::string &error);
    
    cell offset(row_t row, column_t column);
    
    worksheet get_parent();
    const worksheet get_parent() const;

    value &get_value();
    const value &get_value() const;

	void set_value(bool b);
	void set_value(std::int8_t i);
	void set_value(std::int16_t i);
	void set_value(std::int32_t i);
	void set_value(std::int64_t i);
	void set_value(std::uint8_t i);
	void set_value(std::uint16_t i);
	void set_value(std::uint32_t i);
	void set_value(std::uint64_t i);
#ifdef _WIN32
	void set_value(unsigned long i);
#endif
	void set_value(float f);
	void set_value(double d);
	void set_value(long double d);
	void set_value(const char *s);
	void set_value(const std::string &s);
	void set_value(const date &d);
	void set_value(const datetime &d);
	void set_value(const time &t);
	void set_value(const timedelta &t);
	void set_value(const value &v);
    
    cell &operator=(const cell &rhs);
    
    bool operator==(const cell &comparand) const;
    bool operator==(std::nullptr_t) const;

    friend bool operator==(std::nullptr_t, const cell &cell);
    friend bool operator<(cell left, cell right);
    
private:
    friend class worksheet;
    cell(detail::cell_impl *d);
    detail::cell_impl *d_;
};

inline std::ostream &operator<<(std::ostream &stream, const xlnt::cell &cell)
{
    return stream << cell.to_string();
}
    
} // namespace xlnt
