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
#include <xlnt/worksheet/worksheet.hpp>

#include "cell_impl.hpp"
#include "comment_impl.hpp"

namespace {

std::pair<bool, long double> cast_numeric(const std::string &s)
{
    const char *str = s.c_str();
    char *str_end = nullptr;
    auto result = std::strtold(str, &str_end);
    if (str_end != str + s.size()) return { false, 0 };
    return { true, result };
}

std::pair<bool, long double> cast_percentage(const std::string &s)
{
    if (s.back() == '%')
    {
        auto number = cast_numeric(s.substr(0, s.size() - 1));

        if (number.first)
        {
            return { true, number.second / 100 };
        }
    }

    return { false, 0 };
}

std::pair<bool, xlnt::time> cast_time(const std::string &s)
{
    xlnt::time result;

    try
    {
        auto last_colon = s.find_last_of(':');
        if (last_colon == std::string::npos) return { false, result };
        double seconds = std::stod(s.substr(last_colon + 1));
        result.second = static_cast<int>(seconds);
        result.microsecond = static_cast<int>((seconds - static_cast<double>(result.second)) * 1e6);

        auto first_colon = s.find_first_of(':');

        if (first_colon == last_colon)
        {
            auto decimal_pos = s.find('.');
            if (decimal_pos != std::string::npos)
            {
                result.minute = std::stoi(s.substr(0, first_colon));
            }
            else
            {
                result.hour = std::stoi(s.substr(0, first_colon));
                result.minute = result.second;
                result.second = 0;
            }
        }
        else
        {
            result.hour = std::stoi(s.substr(0, first_colon));
            result.minute = std::stoi(s.substr(first_colon + 1, last_colon - first_colon - 1));
        }
    }
    catch (std::invalid_argument)
    {
        return { false, result };
    }

    return { true, result };
}

} // namespace

namespace xlnt {
namespace detail {

cell_impl::cell_impl() : cell_impl(column_t(1), 1)
{
}

cell_impl::cell_impl(column_t column, row_t row) : cell_impl(nullptr, column, row)
{
}

cell_impl::cell_impl(worksheet_impl *parent, column_t column, row_t row)
    : type_(cell::type::null),
      parent_(parent),
      column_(column),
      row_(row),
      value_numeric_(0),
      has_hyperlink_(false),
      is_merged_(false),
      has_style_(false),
      style_id_(0),
      comment_(nullptr)
{
}

cell_impl::cell_impl(const cell_impl &rhs)
{
    *this = rhs;
}

cell_impl &cell_impl::operator=(const cell_impl &rhs)
{
    parent_ = rhs.parent_;
    value_numeric_ = rhs.value_numeric_;
    value_string_ = rhs.value_string_;
    hyperlink_ = rhs.hyperlink_;
    formula_ = rhs.formula_;
    column_ = rhs.column_;
    row_ = rhs.row_;
    is_merged_ = rhs.is_merged_;
    has_hyperlink_ = rhs.has_hyperlink_;
    type_ = rhs.type_;
    style_id_ = rhs.style_id_;
    has_style_ = rhs.has_style_;

    if (rhs.comment_ != nullptr)
    {
        comment_.reset(new comment_impl(*rhs.comment_));
    }

    return *this;
}

cell cell_impl::self()
{
    return xlnt::cell(this);
}

void cell_impl::set_string(const std::string &s, bool guess_types)
{
    value_string_ = s;
    type_ = cell::type::string;

    if (value_string_.size() > 1 && value_string_.front() == '=')
    {
        formula_ = value_string_;
        type_ = cell::type::formula;
        value_string_.clear();
    }
    else if (cell::error_codes().find(s) != cell::error_codes().end())
    {
        type_ = cell::type::error;
    }
    else if (guess_types)
    {
        auto percentage = cast_percentage(s);

        if (percentage.first)
        {
            value_numeric_ = percentage.second;
            type_ = cell::type::numeric;
            self().set_number_format(xlnt::number_format::percentage());
        }
        else
        {
            auto time = cast_time(s);

            if (time.first)
            {
                type_ = cell::type::numeric;
                self().set_number_format(number_format::date_time6());
                value_numeric_ = time.second.to_number();
            }
            else
            {
                auto numeric = cast_numeric(s);

                if (numeric.first)
                {
                    value_numeric_ = numeric.second;
                    type_ = cell::type::numeric;
                }
            }
        }
    }
}

} // namespace detail
} // namespace xlnt
