// Copyright (c) 2014 Thomas Fussell
// Copyright (c) 2010-2014 openpyxl
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
class relationship;
class worksheet;

struct date;
struct datetime;
struct time;
struct timedelta;

namespace detail {    
struct cell_impl;
} // namespace detail
    
class comment
{
public:
    comment(const std::string &type, const std::string &value) : type_(type), value_(value) {}
    std::string get_type() const { return type_; }
    std::string get_value() const { return value_; }
private:
    std::string type_;
    std::string value_;
};

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
    enum class type
    {
        null,
        numeric,
        string,
        formula,
        boolean,
        error
    };

    static const std::unordered_map<std::string, int> ErrorCodes;
    
    static std::string check_string(const std::string &value);
    static std::string check_numeric(const std::string &value);
    static std::string check_error(const std::string &value);
    
    cell();
    cell(worksheet ws, const cell_reference &reference, const std::string &initial_value = std::string());
    
    std::string get_internal_value_string() const;
    long double get_internal_value_numeric() const;
    
    std::string get_column() const;
    row_t get_row() const;
    
    std::string to_string() const;
    
    void set_explicit_value(const std::string &value, type data_type);
    void set_explicit_value(int value, type data_type);
    void set_explicit_value(double value, type data_type);
    void set_explicit_value(const date &value, type data_type);
    void set_explicit_value(const time &value, type data_type);
    void set_explicit_value(const datetime &value, type data_type);
    type data_type_for_value(const std::string &value);
    
    bool is_merged() const;
    void set_merged(bool merged);
    
    relationship get_hyperlink() const;
    void set_hyperlink(const std::string &value);
    bool has_hyperlink() const;
    
    void set_number_format(const std::string &format_code);
    
    bool has_style() const;
    
    style &get_style();
    const style &get_style() const;
    
    type get_data_type() const;
    
    cell_reference get_reference() const;
    
    bool is_date() const;
    
    //std::pair<int, int> get_anchor() const;
    
    comment get_comment() const;
    void set_comment(const comment &comment);
    void clear_comment();

    std::string get_formula() const;
    void set_formula(const std::string &formula);

    std::string get_error() const;
    void set_error(const std::string &error);

    void set_null();
    
    cell offset(row_t row, column_t column);
    
    worksheet get_parent();
    
    cell &operator=(const cell &rhs);
    cell &operator=(bool value);
    cell &operator=(int value);
    cell &operator=(double value);
    cell &operator=(long int value);
    cell &operator=(long long value);
    cell &operator=(long double value);
    cell &operator=(const std::string &value);
    cell &operator=(const char *value);
    cell &operator=(const date &value);
    cell &operator=(const time &value);
    cell &operator=(const datetime &value);
    cell &operator=(const timedelta &value);
    
    bool operator==(const cell &comparand) const;
    bool operator==(std::nullptr_t) const;
    bool operator==(bool comparand) const;
    bool operator==(int comparand) const;
    bool operator==(double comparand) const;
    bool operator==(const std::string &comparand) const;
    bool operator==(const char *comparand) const;
    bool operator==(const date &comparand) const;
    bool operator==(const time &comparand) const;
    bool operator==(const datetime &comparand) const;
    bool operator==(const timedelta &comparand) const;

    friend bool operator==(std::nullptr_t, const cell &cell);
    friend bool operator==(bool comparand, const cell &cell);
    friend bool operator==(int comparand, const cell &cell);
    friend bool operator==(double comparand, const cell &cell);
    friend bool operator==(const std::string &comparand, const cell &cell);
    friend bool operator==(const char *comparand, const cell &cell);
    friend bool operator==(const date &comparand, const cell &cell);
    friend bool operator==(const time &comparand, const cell &cell);
    friend bool operator==(const datetime &comparand, const cell &cell);
    
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
